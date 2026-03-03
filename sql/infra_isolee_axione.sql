-- ============================================================
-- Detection des infrastructures isolees pour les SRO Axione
-- Tables: infra_pt_pot, infra_pt_chb, infra_pt_autres, bpe
-- Reference cables: rip_avg_nge.cables
-- ============================================================

CREATE SCHEMA IF NOT EXISTS _controles;

-- ============================================================
-- 1. POTEAUX ISOLES (infra_pt_pot)
-- Classification:
--   - ISOLE        : aucun cable dans un rayon de 2m
--   - PROCHE (<2m) : pas d'intersection mais cable a moins de 2m
--   - OK           : intersecte un cable (exclu du resultat)
-- ============================================================
DROP TABLE IF EXISTS _controles.poteaux_isolees_axione;
CREATE TABLE _controles.poteaux_isolees_axione AS
WITH sro_axione AS (
    SELECT sro
    FROM comac.sro_nge_axione
    WHERE be = 'axione'
),
poteaux AS (
    SELECT p.gid, p.inf_num, p.sro, p.geom
    FROM rip_avg_nge.infra_pt_pot p
    JOIN sro_axione s ON p.sro = s.sro
    WHERE p.geom IS NOT NULL
),
cable_intersect AS (
    SELECT DISTINCT p.gid
    FROM poteaux p
    JOIN rip_avg_nge.cables c ON ST_DWithin(p.geom, c.geom, 0.01)
    WHERE c.geom IS NOT NULL
),
cable_proche AS (
    SELECT DISTINCT ON (p.gid)
        p.gid,
        ST_Distance(p.geom, c.geom) AS distance_cable_m,
        c.gid AS gid_cable,
        c.cb_etiquet
    FROM poteaux p
    JOIN rip_avg_nge.cables c ON ST_DWithin(p.geom, c.geom, 2.0)
    WHERE c.geom IS NOT NULL
      AND p.gid NOT IN (SELECT gid FROM cable_intersect)
    ORDER BY p.gid, ST_Distance(p.geom, c.geom)
)
SELECT
    p.gid,
    p.inf_num,
    p.sro,
    CASE
        WHEN cp.gid IS NOT NULL THEN 'PROCHE'
        ELSE 'ISOLE'
    END AS statut,
    ROUND(cp.distance_cable_m::numeric, 3) AS distance_cable_m,
    cp.gid_cable,
    cp.cb_etiquet AS cable_etiquette,
    p.geom
FROM poteaux p
LEFT JOIN cable_intersect ci ON ci.gid = p.gid
LEFT JOIN cable_proche cp ON cp.gid = p.gid
WHERE ci.gid IS NULL
ORDER BY
    CASE WHEN cp.gid IS NULL THEN 0 ELSE 1 END,
    cp.distance_cable_m NULLS FIRST;

CREATE INDEX ON _controles.poteaux_isolees_axione USING GIST (geom);


-- ============================================================
-- 2. BPE ISOLES (rip_avg_nge.bpe)
-- Meme classification que les poteaux
-- ============================================================
DROP TABLE IF EXISTS _controles.bpe_isolees_axione;
CREATE TABLE _controles.bpe_isolees_axione AS
WITH sro_axione AS (
    SELECT sro
    FROM comac.sro_nge_axione
    WHERE be = 'axione'
),
bpe_points AS (
    SELECT b.gid, b.inf_num, b.noe_type, b.noe_usage, b.sro, b.geom
    FROM rip_avg_nge.bpe b
    JOIN sro_axione s ON b.sro = s.sro
    WHERE b.geom IS NOT NULL
),
cable_intersect AS (
    SELECT DISTINCT b.gid
    FROM bpe_points b
    JOIN rip_avg_nge.cables c ON ST_DWithin(b.geom, c.geom, 0.01)
    WHERE c.geom IS NOT NULL
),
cable_proche AS (
    SELECT DISTINCT ON (b.gid)
        b.gid,
        ST_Distance(b.geom, c.geom) AS distance_cable_m,
        c.gid AS gid_cable,
        c.cb_etiquet
    FROM bpe_points b
    JOIN rip_avg_nge.cables c ON ST_DWithin(b.geom, c.geom, 2.0)
    WHERE c.geom IS NOT NULL
      AND b.gid NOT IN (SELECT gid FROM cable_intersect)
    ORDER BY b.gid, ST_Distance(b.geom, c.geom)
)
SELECT
    b.gid,
    b.inf_num,
    b.noe_type,
    b.noe_usage,
    b.sro,
    CASE
        WHEN cp.gid IS NOT NULL THEN 'PROCHE'
        ELSE 'ISOLE'
    END AS statut,
    ROUND(cp.distance_cable_m::numeric, 3) AS distance_cable_m,
    cp.gid_cable,
    cp.cb_etiquet AS cable_etiquette,
    b.geom
FROM bpe_points b
LEFT JOIN cable_intersect ci ON ci.gid = b.gid
LEFT JOIN cable_proche cp ON cp.gid = b.gid
WHERE ci.gid IS NULL
ORDER BY
    CASE WHEN cp.gid IS NULL THEN 0 ELSE 1 END,
    cp.distance_cable_m NULLS FIRST;

CREATE INDEX ON _controles.bpe_isolees_axione USING GIST (geom);


-- ============================================================
-- 3. CHAMBRES ISOLEES (infra_pt_chb)
-- Meme classification
-- ============================================================
DROP TABLE IF EXISTS _controles.chambres_isolees_axione;
CREATE TABLE _controles.chambres_isolees_axione AS
WITH sro_axione AS (
    SELECT sro
    FROM comac.sro_nge_axione
    WHERE be = 'axione'
),
chambres AS (
    SELECT ch.gid, ch.inf_num, ch.sro, ch.geom
    FROM rip_avg_nge.infra_pt_chb ch
    JOIN sro_axione s ON ch.sro = s.sro
    WHERE ch.geom IS NOT NULL
),
cable_intersect AS (
    SELECT DISTINCT ch.gid
    FROM chambres ch
    JOIN rip_avg_nge.cables c ON ST_DWithin(ch.geom, c.geom, 0.01)
    WHERE c.geom IS NOT NULL
),
cable_proche AS (
    SELECT DISTINCT ON (ch.gid)
        ch.gid,
        ST_Distance(ch.geom, c.geom) AS distance_cable_m,
        c.gid AS gid_cable,
        c.cb_etiquet
    FROM chambres ch
    JOIN rip_avg_nge.cables c ON ST_DWithin(ch.geom, c.geom, 2.0)
    WHERE c.geom IS NOT NULL
      AND ch.gid NOT IN (SELECT gid FROM cable_intersect)
    ORDER BY ch.gid, ST_Distance(ch.geom, c.geom)
)
SELECT
    ch.gid,
    ch.inf_num,
    ch.sro,
    CASE
        WHEN cp.gid IS NOT NULL THEN 'PROCHE'
        ELSE 'ISOLE'
    END AS statut,
    ROUND(cp.distance_cable_m::numeric, 3) AS distance_cable_m,
    cp.gid_cable,
    cp.cb_etiquet AS cable_etiquette,
    ch.geom
FROM chambres ch
LEFT JOIN cable_intersect ci ON ci.gid = ch.gid
LEFT JOIN cable_proche cp ON cp.gid = ch.gid
WHERE ci.gid IS NULL
ORDER BY
    CASE WHEN cp.gid IS NULL THEN 0 ELSE 1 END,
    cp.distance_cable_m NULLS FIRST;

CREATE INDEX ON _controles.chambres_isolees_axione USING GIST (geom);


-- ============================================================
-- 4. RESUME PAR SRO (comptage rapide)
-- Nombre de poteaux/BPE/chambres isoles par SRO Axione
-- ============================================================
DROP TABLE IF EXISTS _controles.resume_poteaux_bpe_chambres_isolees_axione;
CREATE TABLE _controles.resume_poteaux_bpe_chambres_isolees_axione AS
WITH sro_axione AS (
    SELECT sro
    FROM comac.sro_nge_axione
    WHERE be = 'axione'
),
pot_isole AS (
    SELECT p.sro, p.gid
    FROM rip_avg_nge.infra_pt_pot p
    JOIN sro_axione s ON p.sro = s.sro
    WHERE p.geom IS NOT NULL
      AND NOT EXISTS (
          SELECT 1 FROM rip_avg_nge.cables c
          WHERE c.geom IS NOT NULL AND ST_DWithin(p.geom, c.geom, 0.01)
      )
),
bpe_isole AS (
    SELECT b.sro, b.gid
    FROM rip_avg_nge.bpe b
    JOIN sro_axione s ON b.sro = s.sro
    WHERE b.geom IS NOT NULL
      AND NOT EXISTS (
          SELECT 1 FROM rip_avg_nge.cables c
          WHERE c.geom IS NOT NULL AND ST_DWithin(b.geom, c.geom, 0.01)
      )
),
chb_isole AS (
    SELECT ch.sro, ch.gid
    FROM rip_avg_nge.infra_pt_chb ch
    JOIN sro_axione s ON ch.sro = s.sro
    WHERE ch.geom IS NOT NULL
      AND NOT EXISTS (
          SELECT 1 FROM rip_avg_nge.cables c
          WHERE c.geom IS NOT NULL AND ST_DWithin(ch.geom, c.geom, 0.01)
      )
)
SELECT
    s.sro,
    COALESCE(p.nb_pot, 0)  AS poteaux_isoles,
    COALESCE(b.nb_bpe, 0)  AS bpe_isoles,
    COALESCE(ch.nb_chb, 0) AS chambres_isolees,
    COALESCE(p.nb_pot, 0) + COALESCE(b.nb_bpe, 0) + COALESCE(ch.nb_chb, 0) AS total_isoles
FROM sro_axione s
LEFT JOIN (SELECT sro, COUNT(*) AS nb_pot FROM pot_isole GROUP BY sro) p ON p.sro = s.sro
LEFT JOIN (SELECT sro, COUNT(*) AS nb_bpe FROM bpe_isole GROUP BY sro) b ON b.sro = s.sro
LEFT JOIN (SELECT sro, COUNT(*) AS nb_chb FROM chb_isole GROUP BY sro) ch ON ch.sro = s.sro
ORDER BY total_isoles DESC;
