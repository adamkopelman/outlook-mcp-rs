//! Pure (COM-free) logic behind the `user_is_required` / `user_is_optional`
//! event fields (issue #44): whether the calendar's owner is invited as a
//! required or an optional attendee, matching by address (never by display
//! name) and looking through distribution lists, including nested ones.
//!
//! Expanding lists can be expensive (each member is an Outlook call, and a
//! big list may go to the server), so the work is bounded per tool call:
//! - lists nest at most [`MAX_LIST_DEPTH`] levels deep;
//! - at most [`MAX_LIST_MEMBERS_PER_CALL`] members are read in total, shared
//!   by every list of every event in the call;
//! - the COM side adds a time budget (`members_of` returns `None` past it);
//! - each list is expanded at most once per call ([`ListCache`]), and the
//!   expansion is skipped entirely when a direct match already decides.
//!
//! When a bound stops the search before an answer is certain, the role is
//! reported as unknown (`None`) rather than a guess.
//!
//! The COM side (`client.rs`) only reads the raw data, an item's
//! `Recipients` and a list's `AddressEntry.Members`, and hands it to these
//! functions.

use std::collections::{HashMap, HashSet};

use crate::constants::{OL_RECIPIENT_OPTIONAL, OL_RECIPIENT_REQUIRED};

/// How deep nested distribution lists are followed: a list directly on the
/// invite is depth 1, a list inside it depth 2, and so on.
pub const MAX_LIST_DEPTH: usize = 10;
/// Upper bound on list members read during one tool call, across every list
/// and event, so a huge organization-wide list can't stall `list_events`.
pub const MAX_LIST_MEMBERS_PER_CALL: usize = 5000;

/// One entry of an item's `Recipients` collection, or one member of a list.
#[derive(Debug, Clone, PartialEq)]
pub struct Attendee {
    /// `Recipient.Type`: 1 = required, 2 = optional, 3 = resource. Members
    /// of a list carry their list's type.
    pub kind: i32,
    /// `Recipient.Address` / `AddressEntry.Address`: the SMTP address, or
    /// the Exchange legacy DN for an Exchange entry.
    pub address: String,
    /// True for a distribution list (Exchange or personal contact group).
    pub is_list: bool,
}

/// The addresses that identify the calendar's owner (the current user, or
/// the `calendar_of` person): typically the Exchange legacy DN and the
/// primary SMTP address. Compared case-insensitively and trimmed.
#[derive(Debug, Clone, Default)]
pub struct Identity {
    addresses: HashSet<String>,
}

impl Identity {
    pub fn new<I: IntoIterator<Item = S>, S: AsRef<str>>(addresses: I) -> Self {
        Self {
            addresses: addresses
                .into_iter()
                .map(|a| a.as_ref().trim().to_lowercase())
                .filter(|a| !a.is_empty())
                .collect(),
        }
    }

    pub fn is_empty(&self) -> bool {
        self.addresses.is_empty()
    }

    pub fn matches(&self, address: &str) -> bool {
        let a = address.trim().to_lowercase();
        !a.is_empty() && self.addresses.contains(&a)
    }
}

/// `user_is_required` / `user_is_optional` for one event. At most one is
/// `Some(true)`. Both `Some(false)`: the owner is certainly not among the
/// attendees (e.g. a forwarded invite). Both `None`: unknown, because a list
/// couldn't be fully expanded within the bounds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct UserRole {
    pub user_is_required: Option<bool>,
    pub user_is_optional: Option<bool>,
}

impl UserRole {
    pub const REQUIRED: Self = Self { user_is_required: Some(true), user_is_optional: Some(false) };
    pub const OPTIONAL: Self = Self { user_is_required: Some(false), user_is_optional: Some(true) };
    pub const NOT_INVITED: Self = Self { user_is_required: Some(false), user_is_optional: Some(false) };
    pub const UNKNOWN: Self = Self { user_is_required: None, user_is_optional: None };
}

/// Per-call state: which lists contain the owner (so a week of meetings sent
/// to the same list expands it once) and the members still allowed to be
/// read. Lives for one tool call.
#[derive(Debug)]
pub struct ListCache {
    contains: HashMap<String, bool>,
    members_left: usize,
}

impl Default for ListCache {
    fn default() -> Self {
        Self::with_budget(MAX_LIST_MEMBERS_PER_CALL)
    }
}

impl ListCache {
    pub fn with_budget(members: usize) -> Self {
        Self { contains: HashMap::new(), members_left: members }
    }

    /// Members that may still be read in this call.
    pub fn members_left(&self) -> usize {
        self.members_left
    }
}

/// Whether a list contains the owner, or whether the bounds stopped the
/// search first.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Membership {
    Yes,
    No,
    Unknown,
}

/// The owner's role on one event.
///
/// - A personal appointment (no attendees) or a meeting the owner organizes
///   counts as required.
/// - Otherwise a direct match on a required (optional) recipient makes the
///   owner required (optional), as does membership of a required (optional)
///   distribution list. Required wins if both apply.
/// - Lists are only expanded when the direct matches leave the answer open.
///
/// `members_of(list_address)` returns a list's members, or `None` if it
/// can't be expanded (a dynamic list, no permission, or the caller's time
/// budget ran out); that makes the answer unknown unless something else
/// decides it.
pub fn user_role(
    is_meeting: bool,
    organizer_is_owner: bool,
    attendees: &[Attendee],
    owner: &Identity,
    cache: &mut ListCache,
    mut members_of: impl FnMut(&str) -> Option<Vec<Attendee>>,
) -> UserRole {
    if !is_meeting || organizer_is_owner {
        return UserRole::REQUIRED;
    }
    let direct = |kind: i32| {
        attendees.iter().any(|a| a.kind == kind && !a.is_list && owner.matches(&a.address))
    };
    if direct(OL_RECIPIENT_REQUIRED) {
        return UserRole::REQUIRED;
    }
    let lists = |kind: i32| attendees.iter().filter(move |a| a.kind == kind && a.is_list);
    // Required lists first: membership there decides even if the owner is
    // also optional.
    let mut required_unknown = false;
    for list in lists(OL_RECIPIENT_REQUIRED) {
        match list_contains(&list.address, owner, cache, &mut members_of) {
            Membership::Yes => return UserRole::REQUIRED,
            Membership::Unknown => required_unknown = true,
            Membership::No => {}
        }
    }
    let mut optional = direct(OL_RECIPIENT_OPTIONAL);
    let mut optional_unknown = false;
    if !optional {
        for list in lists(OL_RECIPIENT_OPTIONAL) {
            match list_contains(&list.address, owner, cache, &mut members_of) {
                Membership::Yes => {
                    optional = true;
                    break;
                }
                Membership::Unknown => optional_unknown = true,
                Membership::No => {}
            }
        }
    }
    match (required_unknown, optional, optional_unknown) {
        (true, _, _) => UserRole::UNKNOWN,
        (false, true, _) => UserRole::OPTIONAL,
        (false, false, true) => UserRole::UNKNOWN,
        (false, false, false) => UserRole::NOT_INVITED,
    }
}

/// Whether the list at `address` contains the owner, directly or through
/// nested lists, within [`MAX_LIST_DEPTH`] and the call's member budget.
/// Definite answers are cached per list address; a list being expanded is
/// never re-entered, so a cycle of lists terminates.
pub fn list_contains(
    address: &str,
    owner: &Identity,
    cache: &mut ListCache,
    members_of: &mut impl FnMut(&str) -> Option<Vec<Attendee>>,
) -> Membership {
    let mut visiting = HashSet::new();
    contains_inner(address, owner, cache, members_of, 1, &mut visiting).0
}

/// The membership, and whether it may be cached. A "no" reached while a
/// cycle cut the search short (some member list was already being searched
/// higher up) only holds within that search, so it isn't cached.
fn contains_inner(
    address: &str,
    owner: &Identity,
    cache: &mut ListCache,
    members_of: &mut impl FnMut(&str) -> Option<Vec<Attendee>>,
    depth: usize,
    visiting: &mut HashSet<String>,
) -> (Membership, bool) {
    let key = address.trim().to_lowercase();
    if let Some(&hit) = cache.contains.get(&key) {
        return (if hit { Membership::Yes } else { Membership::No }, true);
    }
    if !visiting.insert(key.clone()) {
        return (Membership::No, false); // Already being searched higher up.
    }
    let (result, cacheable) = expand(address, owner, cache, members_of, depth, visiting);
    visiting.remove(&key);
    match result {
        Membership::Yes => { cache.contains.insert(key, true); }
        Membership::No if cacheable => { cache.contains.insert(key, false); }
        _ => {}
    }
    (result, cacheable)
}

fn expand(
    address: &str,
    owner: &Identity,
    cache: &mut ListCache,
    members_of: &mut impl FnMut(&str) -> Option<Vec<Attendee>>,
    depth: usize,
    visiting: &mut HashSet<String>,
) -> (Membership, bool) {
    if depth > MAX_LIST_DEPTH || cache.members_left == 0 {
        return (Membership::Unknown, false);
    }
    let Some(members) = members_of(address) else { return (Membership::Unknown, false) };
    let mut complete = true;
    let mut cacheable = true;
    // Direct members first: the cheap answer for the common case.
    for m in &members {
        if cache.members_left == 0 {
            complete = false;
            break;
        }
        cache.members_left -= 1;
        if !m.is_list && owner.matches(&m.address) {
            return (Membership::Yes, true);
        }
    }
    for m in members.iter().filter(|m| m.is_list) {
        let (found, ok) = contains_inner(&m.address, owner, cache, members_of, depth + 1, visiting);
        cacheable &= ok;
        match found {
            Membership::Yes => return (Membership::Yes, true),
            Membership::Unknown => complete = false,
            Membership::No => {}
        }
    }
    if complete { (Membership::No, cacheable) } else { (Membership::Unknown, false) }
}

/// How many attendees an Outlook `RequiredAttendees` / `OptionalAttendees`
/// string names (`"; "`-separated display names).
pub fn attendee_count(attendees: &str) -> usize {
    attendees.split(';').filter(|s| !s.trim().is_empty()).count()
}

#[cfg(test)]
mod tests {
    use super::*;

    const ME: &str = "/o=ExchangeLabs/ou=Exchange/cn=Recipients/cn=me";

    fn me() -> Identity {
        Identity::new([ME, "Me@Example.com"])
    }

    fn person(kind: i32, address: &str) -> Attendee {
        Attendee { kind, address: address.into(), is_list: false }
    }

    fn list(kind: i32, address: &str) -> Attendee {
        Attendee { kind, address: address.into(), is_list: true }
    }

    fn no_lists(_: &str) -> Option<Vec<Attendee>> {
        panic!("no list should be expanded")
    }

    fn role(attendees: &[Attendee], lists: &[(&str, Vec<Attendee>)]) -> UserRole {
        let mut cache = ListCache::default();
        user_role(true, false, attendees, &me(), &mut cache, |a| {
            lists.iter().find(|(k, _)| *k == a).map(|(_, m)| m.clone())
        })
    }

    /// A chain of `n` nested lists l1 -> l2 -> ... -> ln, with the owner in ln.
    fn chain(n: usize) -> Vec<(String, Vec<Attendee>)> {
        (1..=n)
            .map(|i| {
                let member = if i == n { person(1, ME) } else { list(1, &format!("l{}", i + 1)) };
                (format!("l{i}"), vec![member])
            })
            .collect()
    }

    fn role_in(attendees: &[Attendee], lists: &[(String, Vec<Attendee>)]) -> UserRole {
        let mut cache = ListCache::default();
        user_role(true, false, attendees, &me(), &mut cache, |a| {
            lists.iter().find(|(k, _)| k == a).map(|(_, m)| m.clone())
        })
    }

    #[test]
    fn personal_appointment_and_own_meeting_are_required() {
        let mut cache = ListCache::default();
        assert_eq!(user_role(false, false, &[], &me(), &mut cache, no_lists), UserRole::REQUIRED);
        let others = [person(1, "a@x.com")];
        assert_eq!(user_role(true, true, &others, &me(), &mut cache, no_lists), UserRole::REQUIRED);
    }

    #[test]
    fn direct_match_by_address_caselessly() {
        assert_eq!(role(&[person(1, "a@x.com"), person(1, "ME@example.com")], &[]), UserRole::REQUIRED);
        assert_eq!(role(&[person(1, "a@x.com"), person(2, &ME.to_uppercase())], &[]), UserRole::OPTIONAL);
        assert_eq!(role(&[person(1, "a@x.com"), person(3, "me@example.com")], &[]), UserRole::NOT_INVITED);
        assert_eq!(role(&[person(1, "Me")], &[]), UserRole::NOT_INVITED, "display names never match");
    }

    #[test]
    fn required_wins_over_optional() {
        let a = [person(2, "me@example.com"), person(1, ME)];
        assert_eq!(role(&a, &[]), UserRole::REQUIRED);
        let a = [person(2, "me@example.com"), list(1, "team@x.com")];
        assert_eq!(role(&a, &[("team@x.com", vec![person(1, ME)])]), UserRole::REQUIRED);
    }

    #[test]
    fn membership_through_a_list_takes_the_list_type() {
        let lists = [("team@x.com", vec![person(1, "a@x.com"), person(1, "me@example.com")])];
        assert_eq!(role(&[list(1, "team@x.com")], &lists), UserRole::REQUIRED);
        assert_eq!(role(&[list(2, "team@x.com")], &lists), UserRole::OPTIONAL);
    }

    #[test]
    fn an_unexpandable_list_makes_the_answer_unknown_unless_something_decides() {
        let lists = [("team@x.com", vec![person(1, "a@x.com")])];
        assert_eq!(role(&[list(2, "dynamic@x.com")], &lists), UserRole::UNKNOWN);
        assert_eq!(role(&[list(1, "dynamic@x.com"), person(2, ME)], &lists), UserRole::UNKNOWN,
            "required is open, so optional can't be confirmed");
        assert_eq!(role(&[list(2, "dynamic@x.com"), person(1, ME)], &lists), UserRole::REQUIRED);
        assert_eq!(role(&[list(1, "team@x.com"), list(2, "dynamic@x.com"), person(2, ME)], &lists),
            UserRole::OPTIONAL);
    }

    #[test]
    fn nested_lists_are_followed_to_depth_ten() {
        assert_eq!(MAX_LIST_DEPTH, 10);
        assert_eq!(role_in(&[list(1, "l1")], &chain(MAX_LIST_DEPTH)), UserRole::REQUIRED);
        assert_eq!(role_in(&[list(1, "l1")], &chain(MAX_LIST_DEPTH + 1)), UserRole::UNKNOWN,
            "past the depth bound the answer is unknown, not 'no'");
    }

    #[test]
    fn list_cycles_terminate() {
        let lists = [("a", vec![list(1, "b")]), ("b", vec![list(1, "a")])];
        assert_eq!(role(&[list(1, "a")], &lists), UserRole::NOT_INVITED);
    }

    #[test]
    fn a_no_cut_short_by_a_cycle_is_not_cached() {
        // a -> [b, c(owner)], b -> [a]: searching a visits b first, where the
        // cycle back to a stops; b must not be remembered as "no", because b
        // does reach the owner through a -> c.
        let lists = [
            ("a", vec![list(1, "b"), list(1, "c")]),
            ("b", vec![list(1, "a")]),
            ("c", vec![person(1, ME)]),
        ];
        let mut cache = ListCache::default();
        let mut members = |a: &str| lists.iter().find(|(k, _)| *k == a).map(|(_, m)| m.clone());
        assert_eq!(list_contains("a", &me(), &mut cache, &mut members), Membership::Yes);
        assert_eq!(list_contains("b", &me(), &mut cache, &mut members), Membership::Yes);
    }

    #[test]
    fn a_list_is_expanded_once_per_call() {
        let mut calls = 0;
        let mut cache = ListCache::default();
        let mut members = |_: &str| {
            calls += 1;
            Some(vec![person(1, ME)])
        };
        for _ in 0..3 {
            let r = user_role(true, false, &[list(2, "Team@x.com")], &me(), &mut cache, &mut members);
            assert_eq!(r, UserRole::OPTIONAL);
        }
        let r = user_role(true, false, &[list(1, "team@X.com")], &me(), &mut cache, &mut members);
        assert_eq!(r, UserRole::REQUIRED);
        assert_eq!(calls, 1);
    }

    #[test]
    fn a_direct_required_match_skips_list_expansion() {
        let a = [person(1, ME), list(1, "team@x.com"), list(2, "team@x.com")];
        let mut cache = ListCache::default();
        assert_eq!(user_role(true, false, &a, &me(), &mut cache, no_lists), UserRole::REQUIRED);
    }

    #[test]
    fn the_member_budget_is_shared_by_the_whole_call() {
        let mut cache = ListCache::with_budget(5);
        let mut members = |a: &str| {
            Some((0..3).map(|i| person(1, &format!("{a}-{i}@x.com"))).collect::<Vec<_>>())
        };
        // First list: 3 members read, owner not there.
        assert_eq!(list_contains("a", &me(), &mut cache, &mut members), Membership::No);
        assert_eq!(cache.members_left(), 2);
        // Second list: only 2 of 3 can be read, so the answer is unknown and not cached.
        assert_eq!(list_contains("b", &me(), &mut cache, &mut members), Membership::Unknown);
        assert_eq!(cache.members_left(), 0);
        assert!(!cache.contains.contains_key("b"));
        // Out of budget: nothing more is expanded.
        let mut never = |_: &str| -> Option<Vec<Attendee>> { panic!("budget is spent") };
        assert_eq!(list_contains("c", &me(), &mut cache, &mut never), Membership::Unknown);
        // A cached answer is still free.
        assert_eq!(list_contains("A", &me(), &mut cache, &mut never), Membership::No);
    }

    #[test]
    fn empty_identity_matches_nothing() {
        let nobody = Identity::new(["", "  "]);
        assert!(nobody.is_empty());
        assert!(!nobody.matches(""));
    }

    #[test]
    fn counts_attendees_in_outlook_strings() {
        assert_eq!(attendee_count(""), 0);
        assert_eq!(attendee_count("Ada Lovelace"), 1);
        assert_eq!(attendee_count("Ada Lovelace; Bob Smith; "), 2);
    }
}
