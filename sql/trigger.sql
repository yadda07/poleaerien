-- Trigger: trg_before_insert_pt_ac

-- DROP TRIGGER IF EXISTS trg_before_insert_pt_ac ON rip_avg_nge.infra_pt_pot;

CREATE OR REPLACE TRIGGER trg_before_insert_pt_ac
    BEFORE INSERT OR UPDATE 
    ON rip_avg_nge.infra_pt_pot
    FOR EACH ROW
    WHEN (new.inf_type::text = 'POT-AC'::text)
    EXECUTE FUNCTION rip_avg_nge.insert_inf_num_pt_ac();