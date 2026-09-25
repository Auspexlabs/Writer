# Design Doc: Real-Time Collaborative Editing

| Status | In review | Author | Casey Kim | Updated | 2026-09-25 |
| --- | --- | --- | --- | --- | --- |

## Background

Today the last save wins when several people edit one document. That produces about 30 conflicts a week, each merged by hand. The goal: everyone editing the same document sees changes live, merged automatically and attributable.

## Goals and non-goals

**Goals**
- 100 people editing one document with P95 operation latency under 300 ms
- Editing continues offline and merges on reconnect
- Every change is attributed and can be rolled back

**Non-goals**
- No transactions across documents
- Comments are not synced live in the first release

## Design

Operations merge on the client with a CRDT (Yjs); the server only broadcasts and stores snapshots:

```
client A ──┐                       ┌── client B
           ├─ WebSocket ─ sync service ─┤
client C ──┘          │                └── client D
                      ▼
            snapshot store (every 5 min / 500 ops)
```

- The client holds the whole document; offline edits queue locally
- The sync service is stateless and scales out; rooms are sharded by document id
- Recovery replays the log over the last snapshot; logs are kept for 30 days

## API

| Endpoint | Method | Purpose |
| --- | --- | --- |
| `/ws/doc/{id}` | WebSocket | Join the room; two-way updates |
| `/api/doc/{id}/snapshot` | GET | Latest snapshot with its version |
| `/api/doc/{id}/history?from=` | GET | Operations after a version |
| `/api/doc/{id}/restore` | POST | Roll back to a version |

## Data model

```json
{
  "docId": "d_8f3a",
  "version": 1284,
  "snapshot": "<binary>",
  "ops": [{ "seq": 1285, "user": "u_12", "at": "2026-09-25T08:12:03Z", "delta": "<binary>" }]
}
```

## Risks

| Risk | Impact | Mitigation |
| --- | --- | --- |
| Large documents load slowly at first | Poor experience | Chunked loading and incremental rendering |
| The sync service is a single point of failure | Nobody can edit | Two active instances; clients reconnect on their own |
| History grows without bound | Storage cost | Periodic snapshot compaction; logs expire after 30 days |

## Milestones

| Phase | Scope | Date |
| --- | --- | --- |
| M1 | Prototype: two clients syncing, nothing stored | October 15 |
| M2 | Persistence, offline queue, permissions | November 10 |
| M3 | 5% of users for a week | November 24 |
| M4 | Everyone | December 8 |

## Open questions

- Compress operations before M2?
- How do we show attribution for anonymous collaborators?
