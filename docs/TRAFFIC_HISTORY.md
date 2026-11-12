# Traffic history

GitHub keeps repository traffic for 14 days only.
`.github/workflows/daily-traffic.yml` keeps a longer record (#165).

## What the workflow does

- **Schedule.** It runs every day at 05:23 UTC and on manual dispatch, from
  the default branch only.
- **Data.** It reads the traffic API: views, clones, referrers and popular
  paths, plus the repository's star, fork, watcher and open-issue counts.
- **History.** It appends the day's snapshot to the CSV files under `data/` on
  the orphan `traffic-history` branch, which the first run creates:
  - `traffic.csv`
  - `traffic_daily.csv`
  - `referrers.csv`
  - `popular_paths.csv`
- **Badge endpoints.** It writes three shields.io endpoint files under
  `data/badges/`, for the README traffic badges:

  | File | Badge | Value |
  |---|---|---|
  | `clones.json` | clones | Total clones recorded since tracking began. |
  | `visitors.json` | visitors (14d) | Unique visitors over GitHub's rolling 14-day window, from today's snapshot. |
  | `visits.json` | visits | Daily unique visitors summed since tracking began; a visitor counts once per day. |

The `traffic-history` branch is public. It holds only the aggregate counts
GitHub already shows to repository admins, with nothing that identifies a
visitor.

## Security boundary

This is analytics, not CI. The job runs inline collection scripts, never
repository code, and must never be extended to do so.

- **Workflow token.** It defaults to `contents: read`. The one job receives
  `contents: write`, to push `traffic-history`.
- **Checkout.** It keeps no credentials. The token reaches git only in the
  push, as a masked header on that one command.
- **Traffic API token.** The traffic API is read with a short-lived
  installation token of a dedicated GitHub App, holding
  `Administration: read` on this repository only.
- **Who sees what.**
  - Only **Create the traffic App token** receives the private key.
  - Only **Fetch GitHub traffic API data** receives the installation token.
  - The configuration check sees only whether the key is set.
- **Revocation.** The token is revoked when the job ends. There is no
  personal-token fallback.
- **Retries.** API calls retry up to five times within 300 seconds, with a
  60-second timeout per request. All five responses must succeed and validate
  before anything is pushed.
- **Concurrent writers.** A rejected push is rebuilt on the remote history, up
  to three attempts, and never force-pushed.

`tools/test_daily_traffic.py` runs the embedded scripts against local
fixtures and checks this credential and permission layout.
`tools/workflow_controls.py` allows `contents: write` to this job alone.

## Setup

These steps are done by the maintainer.

1. **The App.**
   - Use a GitHub App with only **Administration: Read-only**, plus the
     mandatory Metadata read, and its webhook disabled.
   - Install it on this repository; selecting only the intended repositories
     is enough.
2. **The `analytics` environment.** In **Settings → Environments**, create
   `analytics`:
   - **Deployment branches and tags**: selected branches, `main` only.
   - **Environment variable** `TRAFFIC_APP_CLIENT_ID`: the App's Client ID.
   - **Environment secret** `TRAFFIC_APP_PRIVATE_KEY`: the complete PEM of a
     private key generated for this repository.

   Use environment entries, not repository secrets. Delete the downloaded key
   once it is saved, and never attach it to an issue, pull request or chat.
3. **First run.** Once the workflow is on `main`, start it from **Actions →
   Daily traffic export → Run workflow** on `main`. Then:
   - check that a commit on `traffic-history` added `data/` and
     `data/badges/`;
   - only then add the README traffic badges.

   The first run succeeded on 2026-11-12, and the README has shown the
   clones, visitors and visits badges since then.

## Failures

- **Missing credentials** fail the configuration check with an error naming
  the `analytics` environment.
- **A failed token step** means a wrong Client ID, a wrong private key, or an
  App not installed on this repository.
- **A 403 or 404 from the traffic API** means the installation or its
  Administration read permission is missing.

Fix the configuration and run the workflow again on `main`.

## Key rotation

Rotate the key every 90 days, at once on suspected exposure, or when someone
other than the maintainer changes the workflow:

1. generate a new key in the App settings;
2. replace `TRAFFIC_APP_PRIVATE_KEY` in `analytics`;
3. run the workflow on `main` and check the new data commit;
4. delete the old key in the App settings, after checking its fingerprint.

A compromised key is revoked at once, without waiting for that check.
