-- ############################################################################################
-- WHICH CRT VERSIONS CALL WHICH ROUTE (owner request, 2026-10-04: "how about tracking the API
-- end-points, to see if it is possible to retire any, if almost no versions uses it any more").
--
-- One row per day, route and CRT version: how many calls, and the last. Counted in memory as the
-- requests arrive (ApiUsageCounter) and added here every few minutes (ApiUsageFlusher), so a busy
-- minute is one write, not one per request. Read by Admin > "API usage" (ApiUsageFlow).
--
--   callDate  - the UTC day.
--   method    - "GET", "POST" ...
--   route     - the route's PATTERN, never a real path ("/api/review/submissions/{submissionId:long}"),
--               so no id, address or account ever lands here - and a route is one row however many
--               submissions it is called for. Case-sensitive, like the routes.
--   version   - the CRT version the User-Agent names ("3.0.0-beta.2"), "(not CRT)" for a request
--               naming none (a browser, an uptime check), "(other)" past the per-day cap.
--   calls     - how many.
--   lastUtc   - the last call that day, UTC.
--
-- NO ADDRESS, NO ACCOUNT, NO USER-AGENT TEXT beyond the version - the table says which versions
-- still use a route, which is all retiring one needs. Kept until a reset (DataResetFlow) empties it;
-- a row a day per route and version in use is small.
--
-- camelCase columns like crt_board_views (0013), the other crt_* usage table.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################

CREATE TABLE crt_api_calls (
    callDate   DATE             NOT NULL,
    method     VARCHAR(10)      NOT NULL,
    route      VARCHAR(200)     NOT NULL COLLATE utf8mb4_bin,
    version    VARCHAR(50)      NOT NULL,
    calls      BIGINT UNSIGNED  NOT NULL,
    lastUtc    DATETIME         NOT NULL,

    PRIMARY KEY (callDate, method, route, version)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
