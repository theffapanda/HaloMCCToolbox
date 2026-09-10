# Ranked SmartMatch rank-21 experiment

## Status

Disabled and removed from the Toolbox runtime on 2026-07-21. `ProxyService` is back to passive SmartMatch capture. This note preserves the findings and the narrow implementation points if the controlled experiment is intentionally resumed later.

## Captured controls

| Playlist | Hopper | Rank | Percentile | SmartMatch wait estimate |
| --- | --- | ---: | ---: | ---: |
| Halo 3 Ranked Team Slayer | `Cascade8PTeamRanked` | 39 | `0.9994376301765442` | 25,879 seconds |
| Reach Hardcore | `Cascade8PTeamRankedLoose` | 9 | `0.6637910604476929` | 2,590 seconds |
| Ranked 2v2 | `Cascade4PTeamRanked` | 21 | `0.726523756980896` | 191 seconds |

The rank-21 tuple used for the test was:

- Level and recommended set point: `21`
- Percentile: `0.726523756980896`

All nine related ticket attributes were changed together:

- `AverageGroupSkillLevel`
- `MinGroupSkillLevel`
- `MaxGroupSkillLevel`
- `RecommendedMatchmakingSetPoint`
- `MinRecommendedMatchmakingSetPoint`
- `MaxRecommendedMatchmakingSetPoint`
- `AverageGroupSkillPercentile`
- `MinGroupSkillPercentile`
- `MaxGroupSkillPercentile`

## Result

The test account's real ranked-2v2 ticket advertised rank 1. The proxy rewrote the outbound ranked ticket to the captured rank-21 tuple. SmartMatch returned HTTP 200 and the completed `CascadeMatchmaking` document preserved the modified attributes for the test account.

The resulting match contained:

- Test account: advertised rank 21 (real rank 1)
- Teammate: rank 23
- Opposing two-player party: minimum rank 2, maximum rank 21, average rank 11, recommended set point 21

This is strong evidence that this ranked hopper used the client-submitted skill tuple when assembling the match. The opposing party's per-player ranks could not be assigned to individual names because MPSD returned the same party aggregate on both party members.

## Previous narrow implementation

The interception lived in `ProxyService.OnBeforeRequestAsync`, immediately after the existing safe SmartMatch request-body capture. It applied only when all of these conditions were true:

- Method was `POST`
- Host ended with `smartmatch.xboxlive.com`
- Path contained `/hoppers/`
- Path contained `Ranked`
- Request had no `Content-Encoding`
- JSON contained `ticketAttributes` and all nine expected skill fields

The request JSON was parsed as a `JsonObject`, the six level/set-point fields were set to `21`, and the three percentile fields were set to `0.726523756980896`. The compact JSON was encoded as UTF-8 and passed to `e.SetRequestBody(...)`; `entry.RequestBody` was also replaced so the saved diagnostic ticket represented what was actually sent.

The first version rewrote only one POST. That was insufficient because MCC renews SmartMatch tickets after the 120-second `giveUpDuration`. The second version kept rewriting every ranked renewal until cancellation. The final test version rewrote every ranked POST for the lifetime of the `ProxyService` instance while leaving social tickets untouched.

## If intentionally restored

Prefer an explicit, visibly labeled developer-only opt-in rather than an always-on field initializer. Keep the host/path/method/content-encoding and nine-field validation above. Log every rewrite and show the active state in the UI. Rebuild and verify that the saved `last-smartmatch-ticket.json` contains the modified tuple before drawing conclusions from a match.
