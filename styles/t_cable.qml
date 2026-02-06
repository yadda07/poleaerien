<!DOCTYPE qgis PUBLIC 'http://mrcc.com/qgis.dtd' 'SYSTEM'>
<qgis version="3.28" styleCategories="AllStyleCategories">
  <renderer-v2 type="singleSymbol" forceraster="0" enableorderby="0">
    <symbols>
      <symbol type="line" name="0" clip_to_extent="1" alpha="1">
        <layer class="SimpleLine" pass="0" enabled="1" locked="0">
          <prop v="0" k="align_dash_pattern"/>
          <prop v="flat" k="capstyle"/>
          <prop v="255,0,0,255" k="line_color"/>
          <prop v="solid" k="line_style"/>
          <prop v="1.2" k="line_width"/>
          <prop v="MM" k="line_width_unit"/>
          <prop v="0" k="offset"/>
          <prop v="MM" k="offset_unit"/>
          <prop v="0" k="tweak_dash_pattern_on_corners"/>
          <prop v="0" k="use_custom_dash"/>
        </layer>
      </symbol>
    </symbols>
  </renderer-v2>
  <labeling type="simple">
    <settings calloutType="simple">
      <text-style fieldName="concat(appui_start, ' → ', appui_end, ' (', cab_capa, ')')" fontFamily="Arial" fontSize="7" fontWeight="50" textColor="255,0,0,255"/>
    </settings>
  </labeling>
</qgis>
