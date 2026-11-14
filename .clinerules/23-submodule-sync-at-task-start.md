# Sync repositories and submodules at task start

Run this refresh at the **start of a new task** — e.g. the first message of a
new chat or a new day — but **not** repeatedly within the same chat/session, and
never over in-progress local work.

- A new task/day → sync:
  1. `git pull` this repository.
  2. `git submodule update --init --recursive` to the recorded pointers.
  3. Advance submodules owned by `andreyputilovmaterial`, or `pinliner` /
     `myers-diff` from `The-city-not-present`, to their latest remote commit
     (checkout the remote default-branch HEAD).
  4. Keep any other external submodule pinned to the pointer recorded in git
     history (do not advance it).
- A repeat message in the same chat → do not re-sync; keep current local state
  and continue the work in progress.
- Always follow the submodule commit-order rule (push a submodule first, and
  record its pointer in the parent only after the commit exists on the
  submodule's remote).
