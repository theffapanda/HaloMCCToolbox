use std::collections::HashSet;
use std::sync::{Arc, RwLock};

pub type PolicyState = Arc<RwLock<HashSet<u32>>>;

pub fn new() -> PolicyState {
    Arc::new(RwLock::new(HashSet::new()))
}

pub fn contains(state: &PolicyState, pid: u32) -> bool {
    state.read().unwrap().contains(&pid)
}

pub fn add(state: &PolicyState, pid: u32) {
    state.write().unwrap().insert(pid);
}

pub fn replace(state: &PolicyState, pids: impl IntoIterator<Item = u32>) {
    let mut guard = state.write().unwrap();
    guard.clear();
    guard.extend(pids);
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replacing_policy_removes_stale_pids() {
        let state = new();
        add(&state, 10);
        add(&state, 20);
        replace(&state, [20, 30]);
        assert!(!contains(&state, 10));
        assert!(contains(&state, 20));
        assert!(contains(&state, 30));
    }
}
