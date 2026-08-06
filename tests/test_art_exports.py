import offipy


def test_offipy_exports_art_api():
    for name in (
        "ART_SCHEMA_VERSION",
        "ART_REPORT_SCHEMA_VERSION",
        "ArtColor",
        "ArtElement",
        "ArtElementRef",
        "ArtFinding",
        "ArtProfile",
        "ArtReport",
        "ArtReportDiff",
        "ArtScene",
        "ArtSlide",
        "ArtSlideReport",
        "ArtTextRun",
        "ArtWarning",
        "DeckQualityReport",
        "DimensionAssessment",
        "analyze_deck",
        "analyze_scene",
        "build_scene",
        "compare_reports",
        "get_profile",
        "merge_scenes",
        "profile_names",
        "render_html",
        "render_markdown",
        "report_to_json",
    ):
        assert hasattr(offipy, name), f"offipy 未导出 {name}"


def test_art_package_exports():
    import offipy.art as art

    for name in (
        "ArtScene",
        "ArtReport",
        "build_scene",
        "merge_scenes",
        "analyze_deck",
        "analyze_scene",
        "get_profile",
        "profile_names",
        "render_html",
        "render_markdown",
        "report_to_json",
    ):
        assert hasattr(art, name), f"offipy.art 未导出 {name}"
