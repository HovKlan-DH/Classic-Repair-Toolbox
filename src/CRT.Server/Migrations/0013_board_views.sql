-- ############################################################################################
-- BOARD VIEWS (owner request, 2026-09-27): "I want it to catch every time a user selects a board
-- ... count only when viewed for +10 seconds" and "a new table that can hold all this data in one
-- table, as I then can see usage of boards in countries".
--
-- ONE ROW PER VIEW. CRT counts a view when a published board has been on screen for ten seconds
-- and sends its views in batches (CRT.Data's BoardViewContract). Each row carries what the launch
-- check-in carries - version, operating system, CPU and country - so a view can be counted by
-- board, country, version or platform with a plain GROUP BY.
--
-- *** NOTHING HERE IDENTIFIES A USER. *** No address and no installation id: the country is looked
-- up from the sender's address as the batch arrives and the address is then forgotten. Two views
-- from one person cannot be linked to each other.
--
-- *** THE COLUMN NAMES FOLLOW crt_update, not this schema's snake_case. *** The table was agreed
-- with the project owner in that form: it sits beside crt_update, crt_statistics and crt_versions
-- (moved into this database by hand, 2026-09-27), and the Fun facts PHP queries all of them - the
-- same names make its queries easy to adapt. That also explains the crt_ prefix.
--
--   viewedUtc      - when the board had been on screen long enough (the user's clock, never later
--                    than the server's, to the second; BoardViewRules.StoredTime).
--   systemId       - "Commodore/C64/250407"; utf8mb4_bin like every other system id (0005).
--   hardwareName / boardName - as CRT's drop-down lists show them, taken from the PUBLISHED main
--                    Excel data file on the server, never from the sender.
--   version        - "CRT 2026.10.0", the form crt_update.version has.
--   osHighlevel / osVersion / cpu - as the launch check-in sends them.
--   countryCode / countryName - NULL when the lookup found nothing.
--   fromBeta       - 1 when CRT was downloading from the BETA source (mostly maintainers checking
--                    their own work), so statistics can leave those out.
--   fromLocalNetwork - 1 when the batch came from the server's own network - the project owner's
--                    machines at home, stored while ServerOptions.CountLocalNetworkBoardViews is
--                    on ("For now I would like my own home usage also to count", 2026-09-27), so
--                    they can be left out or deleted later. Their country is the server's own.
--
-- crt_board_view_batches: the batches already stored, so a batch CRT sends again (its answer was
-- lost) is not counted twice. Kept 60 days, deliberately apart from the views: a batch id links
-- the views it carried, and that link is not kept with them.
--
-- READ 0001_initial.sql's header before editing this. Once applied, this file is history.
-- ############################################################################################
CREATE TABLE crt_board_views (
    id            BIGINT UNSIGNED  NOT NULL AUTO_INCREMENT,
    viewedUtc     DATETIME         NOT NULL,
    systemId      VARCHAR(200)     NOT NULL COLLATE utf8mb4_bin,
    hardwareName  VARCHAR(100)     NOT NULL,
    boardName     VARCHAR(100)     NOT NULL,
    version       VARCHAR(100)     NOT NULL,
    osHighlevel   VARCHAR(20)      NULL,
    osVersion     VARCHAR(200)     NULL,
    cpu           VARCHAR(50)      NULL,
    countryCode   CHAR(2)          NULL,
    countryName   VARCHAR(50)      NULL,
    fromBeta      TINYINT(1)       NOT NULL,
    fromLocalNetwork TINYINT(1)    NOT NULL DEFAULT 0,

    PRIMARY KEY (id),
    KEY ix_crt_board_views_viewed (viewedUtc),
    KEY ix_crt_board_views_system (systemId, viewedUtc),
    KEY ix_crt_board_views_country (countryCode, viewedUtc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE crt_board_view_batches (
    batchId       BINARY(16)       NOT NULL,
    receivedUtc   DATETIME         NOT NULL,

    PRIMARY KEY (batchId),
    KEY ix_crt_board_view_batches_received (receivedUtc)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
