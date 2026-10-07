from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping, Sequence

import matplotlib
import pandas as pd

matplotlib.use("Agg", force=True)
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

from .auction_loader import load_auction_csv
from .auction_metrics import summarize_auction_insights
from .auction_sources import load_cross_platform_auction_csvs
from .chart_builder import ChartBuilder
from .config_loader import ConfigLoader
from .data_loader import detect_latest_complete_quarter, load_csv
from .google_slides_builder import (
    DriveChartAssetStore,
    GoogleSlidesGenerationResult,
    build_period_replacement_requests,
    share_copied_presentation,
)
from .google_slides_templates import TemplateConfig
from .google_workspace import (
    PDF_MIME_TYPE,
    GoogleWorkspaceClient,
    GoogleWorkspaceConfig,
)
from .metrics import build_kpi_summary, prepare_report_data, validate_report_data
from .monthly_google_slides_builder import (
    _read_json,
    _resolve_path,
    _send_batch_updates,
    _table_cell_text_by_id,
    _table_cell_text_requests,
    _table_column_width_requests,
    _table_dimensions_by_id,
    _table_format_requests,
    _write_json,
)
from .narrative_generator import (
    generate_auction_bullets,
    generate_mix_bullets,
    generate_overall_bullets,
    generate_scope_bullets,
    generate_trend_bullets,
)
from .other_campaigns import (
    format_other_top_campaigns_table,
    get_wendy_wu_other_top_campaigns_config,
    load_other_campaign_summary,
)
from .trends_loader import TrendsLoader
from .trends_metrics import summarize_trends
from utils.text_report import _format_mix_table


PROJECT_ROOT = Path(__file__).resolve().parents[1]
WWT_UK_QBR_TEMPLATE_MANIFEST = (
    PROJECT_ROOT
    / "docs"
    / "google_slides_templates"
    / "wendy_wu_uk_qbr_test_template.json"
)
WWT_AUS_QBR_TEMPLATE_MANIFEST = (
    PROJECT_ROOT
    / "docs"
    / "google_slides_templates"
    / "wendy_wu_australia_qbr_test_template.json"
)
WENDY_WU_QBR_TEMPLATE_MANIFESTS = {
    "wendy_wu": WWT_UK_QBR_TEMPLATE_MANIFEST,
    "wendy_wu_australia": WWT_AUS_QBR_TEMPLATE_MANIFEST,
}
WENDY_WU_QBR_BUILDER_NAMES = {
    "wendy_wu": "wendy_wu_uk_qbr_template_manifest",
    "wendy_wu_australia": "wendy_wu_australia_qbr_template_manifest",
}
OBSOLETE_WENDY_WU_QBR_TEMPLATE_IDS = {
    "wendy_wu": {
        "1P5L_zODZ1D81QZK5Z8nuZ41ygON3D8GqTjYEeeDqMec",
    },
    "wendy_wu_australia": {
        "1HukrzM7APAQbPJP9Z_9eyD-RNWa2WNSj6mOQ4dFyMQ8",
        "1aI_bgn-lB2GCJVfjGmsDMX-iLlhRi7HpFs0HfS4un4M",
    },
}

NEUTRAL_DELTA_RGB = {"red": 0.42, "green": 0.42, "blue": 0.42}
BLACK_RGB = {"red": 0.0, "green": 0.0, "blue": 0.0}
WHITE_RGB = {"red": 1.0, "green": 1.0, "blue": 1.0}
POSITIVE_RGB = {"red": 0.03, "green": 0.47, "blue": 0.22}
NEGATIVE_RGB = {"red": 0.78, "green": 0.16, "blue": 0.13}
CARD_METRIC_ORDER = ("Sales Leads", "Cost", "CPL", "CVR", "Clicks", "CTR")
# Revenue is added at generation time as a seventh summary tile (the templates hold six).
REVENUE_TILE_METRIC = "Revenue"
REVENUE_TILE_ACCENT_RGB = {"red": 0.16, "green": 0.36, "blue": 0.6}
REVENUE_TILE_ID_SUFFIX = "_rev"
TILE_VALUE_FONT_PT = 13
TILE_VALUE_WIDTH_SHARE = 0.94
AUCTION_TABLE_ID = "p29_i720"
AUCTION_TITLE = "Non-Brand Auction Insights"
AUCTION_HEADERS = (
    "Domain",
    "Imp. Share",
    "Overlap Rate",
    "Pos. Above",
    "Top of Page",
    "Abs. Top",
    "Outranking",
)
AUCTION_VALUE_COLUMNS = (
    "impression_share",
    "overlap_rate",
    "position_above_rate",
    "top_of_page_rate",
    "absolute_top_of_page_rate",
    "outranking_share",
)
AUCTION_MAX_BODY_ROWS = 8
AUCTION_PREFERRED_SOURCE = "Google Ads"
AUCTION_OWN_DOMAIN = "you"
AUCTION_HIGHLIGHT_RGB = {"red": 0.988, "green": 0.894, "blue": 0.925}
# Template table border grey (#CBD5E1, 0.75pt).
AUCTION_BORDER_RGB = {"red": 0.796, "green": 0.835, "blue": 0.882}
AUCTION_BORDER_WEIGHT_EMU = 9525


SUMMARY_SLIDES: dict[str, dict[str, Any]] = {
    "overall": {
        "label": "Overall",
        "scope": ("overall", None),
        "title_id": "p7_i108",
        "subtitle_id": "p7_i109",
        "footer_id": "p7_i111",
        "value_ids": ("p7_i114", "p7_i119", "p7_i124", "p7_i129", "p7_i134", "p7_i139"),
        "delta_ids": ("p7_i116", "p7_i121", "p7_i126", "p7_i131", "p7_i136", "p7_i141"),
        "bullets_id": "p7_i142",
        "chart_ids": {"cpl_cvr": "p7_i144", "cost_leads": "p7_i106"},
    },
    "brand": {
        "label": "Brand",
        "scope": ("campaign", "Brand"),
        "title_id": "p9_i168",
        "subtitle_id": "p9_i169",
        "footer_id": "p9_i171",
        "value_ids": ("p9_i174", "p9_i179", "p9_i184", "p9_i189", "p9_i194", "p9_i199"),
        "delta_ids": ("p9_i176", "p9_i181", "p9_i186", "p9_i191", "p9_i196", "p9_i201"),
        "table_id": "p9_i202",
        "bullets_id": "p9_i203",
    },
    "generic": {
        "label": "Generic",
        "scope": ("campaign", "Generic"),
        "title_id": "p11_i226",
        "subtitle_id": "p11_i227",
        "footer_id": "p11_i229",
        "value_ids": ("p11_i232", "p11_i237", "p11_i242", "p11_i247", "p11_i252", "p11_i257"),
        "delta_ids": ("p11_i234", "p11_i239", "p11_i244", "p11_i249", "p11_i254", "p11_i259"),
        "table_id": "p11_i260",
        "bullets_id": "p11_i261",
    },
    "performance_max": {
        "label": "Performance Max",
        "scope": ("campaign", "Performance Max"),
        "title_id": "p13_i284",
        "subtitle_id": "p13_i285",
        "footer_id": "p13_i287",
        "value_ids": ("p13_i290", "p13_i295", "p13_i300", "p13_i305", "p13_i310", "p13_i315"),
        "delta_ids": ("p13_i292", "p13_i297", "p13_i302", "p13_i307", "p13_i312", "p13_i317"),
        "table_id": "p13_i318",
        "bullets_id": "p13_i319",
    },
    "demand_gen": {
        "label": "Demand Gen",
        "scope": ("campaign", "Demand Gen"),
        "title_id": "p15_i342",
        "subtitle_id": "p15_i343",
        "footer_id": "p15_i345",
        "value_ids": ("p15_i348", "p15_i353", "p15_i358", "p15_i363", "p15_i368", "p15_i373"),
        "delta_ids": ("p15_i350", "p15_i355", "p15_i360", "p15_i365", "p15_i370", "p15_i375"),
        "table_id": "p15_i376",
        "bullets_id": "p15_i377",
    },
    "china": {
        "label": "China",
        "scope": ("destination", "China"),
        "title_id": "p18_i410",
        "subtitle_id": "p18_i411",
        "footer_id": "p18_i413",
        "value_ids": ("p18_i416", "p18_i421", "p18_i426", "p18_i431", "p18_i436", "p18_i441"),
        "delta_ids": ("p18_i418", "p18_i423", "p18_i428", "p18_i433", "p18_i438", "p18_i443"),
        "table_id": "p18_i445",
        "bullets_id": "p18_i444",
        "chart_ids": {"campaign_mix": "p18_i447"},
    },
    "japan": {
        "label": "Japan",
        "scope": ("destination", "Japan"),
        "title_id": "p20_i469",
        "subtitle_id": "p20_i470",
        "footer_id": "p20_i472",
        "value_ids": ("p20_i475", "p20_i480", "p20_i485", "p20_i490", "p20_i495", "p20_i500"),
        "delta_ids": ("p20_i477", "p20_i482", "p20_i487", "p20_i492", "p20_i497", "p20_i502"),
        "table_id": "p20_i504",
        "bullets_id": "p20_i503",
        "chart_ids": {"campaign_mix": "p20_i506"},
    },
    "se_asia": {
        "label": "SE Asia",
        "scope": ("destination", "SE Asia"),
        "title_id": "p22_i528",
        "subtitle_id": "p22_i529",
        "footer_id": "p22_i532",
        "value_ids": ("p22_i535", "p22_i540", "p22_i545", "p22_i550", "p22_i555", "p22_i560"),
        "delta_ids": ("p22_i537", "p22_i542", "p22_i547", "p22_i552", "p22_i557", "p22_i562"),
        "table_id": "p22_i564",
        "bullets_id": "p22_i563",
        "chart_ids": {"campaign_mix": "p22_i565"},
    },
    "india": {
        "label": "India",
        "scope": ("destination", "India"),
        "title_id": "p24_i587",
        "subtitle_id": "p24_i588",
        "footer_id": "p24_i590",
        "value_ids": ("p24_i593", "p24_i598", "p24_i603", "p24_i608", "p24_i613", "p24_i618"),
        "delta_ids": ("p24_i595", "p24_i600", "p24_i605", "p24_i610", "p24_i615", "p24_i620"),
        "table_id": "p24_i622",
        "bullets_id": "p24_i621",
        "chart_ids": {"campaign_mix": "p24_i623"},
    },
    "central_asia_mongolia": {
        "label": "Central Asia & Mongolia",
        "scope": ("destination", "Central Asia & Mongolia"),
        "title_id": "SLIDES_API1312704722_3",
        "subtitle_id": "SLIDES_API1312704722_4",
        "footer_id": "SLIDES_API1312704722_6",
        "value_ids": (
            "SLIDES_API1312704722_9",
            "SLIDES_API1312704722_14",
            "SLIDES_API1312704722_19",
            "SLIDES_API1312704722_24",
            "SLIDES_API1312704722_29",
            "SLIDES_API1312704722_34",
        ),
        "delta_ids": (
            "SLIDES_API1312704722_11",
            "SLIDES_API1312704722_16",
            "SLIDES_API1312704722_21",
            "SLIDES_API1312704722_26",
            "SLIDES_API1312704722_31",
            "SLIDES_API1312704722_36",
        ),
        "table_id": "SLIDES_API1312704722_38",
        "bullets_id": "SLIDES_API1312704722_37",
        "chart_ids": {"campaign_mix": "SLIDES_API1312704722_39"},
    },
    "other": {
        "label": "Other",
        "scope": ("destination", "Other"),
        "title_id": "p26_i646",
        "subtitle_id": "p26_i647",
        "footer_id": "p26_i649",
        "value_ids": ("p26_i652", "p26_i655", "p26_i658", "p26_i661", "p26_i664", "p26_i667"),
        "delta_ids": ("p26_i670", "p26_i671", "p26_i672", "p26_i673", "p26_i674", "p26_i675"),
        "table_id": "p26_i682",
        "bullets_id": "p26_i669",
    },
}

MONTHLY_TREND_SLIDES: dict[str, dict[str, Any]] = {
    "brand": {"scope": ("campaign", "Brand"), "chart_ids": {"cpl_cvr": "p10_i217", "cost_leads": "p10_i216"}},
    "generic": {"scope": ("campaign", "Generic"), "chart_ids": {"cpl_cvr": "p12_i274", "cost_leads": "p12_i275"}},
    "performance_max": {"scope": ("campaign", "Performance Max"), "chart_ids": {"cpl_cvr": "p14_i332", "cost_leads": "p14_i333"}},
    "demand_gen": {"scope": ("campaign", "Demand Gen"), "chart_ids": {"cpl_cvr": "p16_i391", "cost_leads": "p16_i390"}},
    "china": {"scope": ("destination", "China"), "chart_ids": {"cpl_cvr": "p19_i459", "cost_leads": "p19_i460"}},
    "japan": {"scope": ("destination", "Japan"), "chart_ids": {"cpl_cvr": "p21_i518", "cost_leads": "p21_i519"}},
    "se_asia": {"scope": ("destination", "SE Asia"), "chart_ids": {"cpl_cvr": "p23_i578", "cost_leads": "p23_i577"}},
    "india": {"scope": ("destination", "India"), "chart_ids": {"cpl_cvr": "p25_i637", "cost_leads": "p25_i636"}},
    "central_asia_mongolia": {
        "scope": ("destination", "Central Asia & Mongolia"),
        "chart_ids": {"cpl_cvr": "SLIDES_API1312704722_52", "cost_leads": "SLIDES_API1312704722_51"},
    },
}

TREND_SLIDES: dict[str, dict[str, Any]] = {
    "brand": {
        "label": "Brand",
        "chart_title": "Brand Search Interest YTD",
        "terms_kind": "brand",
        "chart_id": "p3_i57",
        "bullets_id": "p3_i59",
        "source_id": "p3_i58",
    },
    "japan": {
        "label": "Japan",
        "chart_title": "Japan Search Demand YTD",
        "terms_kind": "destination",
        "destination": "Japan",
        "chart_id": "p4_i72",
        "bullets_id": "p4_i74",
        "source_id": "p4_i73",
    },
    "china": {
        "label": "China",
        "chart_title": "China Search Demand YTD",
        "terms_kind": "destination",
        "destination": "China",
        "chart_id": "p5_i87",
        "bullets_id": "p5_i89",
        "source_id": "p5_i88",
    },
}


def generate_wendy_wu_qbr_google_slides(
    *,
    client_id: str,
    client_name: str,
    request_dir: str | Path,
    report_artifacts_path: str | Path,
    template: TemplateConfig,
    workspace_config: GoogleWorkspaceConfig,
    google_client: GoogleWorkspaceClient | None = None,
    export_pdf: bool = True,
) -> GoogleSlidesGenerationResult:
    request_path = Path(request_dir)
    outputs_dir = request_path / "outputs"
    outputs_dir.mkdir(parents=True, exist_ok=True)
    artifact_path = Path(report_artifacts_path)
    manifest_path = outputs_dir / "google_slides_generation_manifest.json"
    template_manifest_path = _template_manifest_path(client_id)
    template_manifest = _read_json(template_manifest_path)
    warnings: list[str] = []
    effective_template_id = _effective_template_id(
        client_id, template, template_manifest, warnings
    )

    base_manifest: dict[str, Any] = {
        "schema_version": "1.0",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "client_id": client_id,
        "client_name": client_name,
        "report_mode": "quarterly",
        "builder": _builder_name(client_id),
        "report_artifacts": str(artifact_path),
        "template_id": effective_template_id,
        "configured_template_id": template.template_id,
        "template_key": template.key,
        "template_manifest": str(template_manifest_path),
        "chart_assets": [],
        "permission_cleanup": [],
        "output_sharing": [],
        "warnings": warnings,
    }

    artifact = _read_json(artifact_path)
    client = google_client or GoogleWorkspaceClient(workspace_config)
    asset_store: DriveChartAssetStore | None = None
    copied_id: str | None = None
    copied_url: str | None = None
    qa_pdf_path: Path | None = None
    output_sharing: list[dict[str, Any]] = []
    status = "success"
    message = f"Native {client_name} QBR Google Slides deck generated."
    batch_update_request_count = 0
    payload: dict[str, Any] | None = None

    try:
        payload = build_wendy_wu_qbr_slides_payload(
            request_dir=request_path,
            artifact=artifact,
            client_id=client_id,
            template_manifest=template_manifest,
        )
        warnings.extend(payload.get("warnings") or [])
        copied = client.copy_file(
            effective_template_id,
            _output_deck_title(client_name, payload["period"]["label"]),
            workspace_config.output_folder_id,
        )
        copied_id = str(copied["id"])
        copied_url = f"https://docs.google.com/presentation/d/{copied_id}/edit"
        output_sharing = share_copied_presentation(
            client, copied_id, effective_template_id, warnings
        )
        presentation = client.get_presentation(copied_id)
        asset_store = DriveChartAssetStore(
            client, str(workspace_config.asset_folder_id)
        )

        requests_body: list[dict[str, Any]] = []
        requests_body.extend(build_period_replacement_requests(payload["artifact"]))
        requests_body.extend(_build_scalar_text_requests(payload["shape_text"], presentation))
        requests_body.extend(
            _revenue_tile_requests(presentation, payload.get("revenue_tiles") or {})
        )
        table_dimensions = _table_dimensions_by_id(presentation)
        table_cell_text = _table_cell_text_by_id(presentation)
        table_widths = _table_widths_by_id(presentation)
        for table_id, table_payload in payload["tables"].items():
            requests_body.extend(
                _replace_existing_table_exact_requests(
                    table_id=table_id,
                    values=table_payload["values"],
                    existing_rows=table_dimensions.get(table_id, (0, 0))[0],
                    existing_columns=table_dimensions.get(table_id, (0, 0))[1],
                    existing_cell_text=table_cell_text.get(table_id, {}),
                    column_widths=table_widths.get(table_id, []),
                )
            )
            requests_body.extend(
                _table_style_requests(table_id, table_payload)
            )

        uploaded_assets = _upload_chart_assets(asset_store, payload["charts"])
        requests_body.extend(_build_chart_requests(uploaded_assets))
        requests_body.extend(
            _build_template_white_text_style_requests(
                presentation, payload["template_manifest"]
            )
        )
        requests_body.extend(_build_cost_delta_style_requests(presentation))

        batch_update_request_count = len(requests_body)
        _send_batch_updates(client, copied_id, requests_body)

        if export_pdf:
            try:
                qa_pdf_path = client.export_file(
                    copied_id, PDF_MIME_TYPE, outputs_dir / "google_slides_qa.pdf"
                )
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"QA PDF export failed: {exc}")
    except Exception as exc:  # noqa: BLE001
        status = "failed"
        message = f"Native {client_name} QBR Google Slides generation failed: {exc}"
    finally:
        cleanup_records = (
            asset_store.cleanup_public_permissions() if asset_store else []
        )
        manifest = {
            **base_manifest,
            "status": status,
            "message": message,
            "copied_presentation_id": copied_id,
            "google_slides_url": copied_url,
            "batch_update_request_count": batch_update_request_count,
            "manual_inputs": payload.get("manual_inputs") if payload else None,
            "automated_sources": payload.get("automated_sources") if payload else None,
            "currency": payload.get("currency") if payload else None,
            "chart_assets": (
                [asset.to_manifest() for asset in asset_store.assets]
                if asset_store
                else []
            ),
            "permission_cleanup": cleanup_records,
            "output_sharing": output_sharing,
            "qa_pdf_path": str(qa_pdf_path) if qa_pdf_path else None,
            "warnings": warnings,
        }
        _write_json(manifest_path, manifest)

    return GoogleSlidesGenerationResult(
        enabled=True,
        status=status,
        message=message,
        google_slides_url=copied_url,
        presentation_id=copied_id,
        manifest_path=manifest_path,
        artifact_path=artifact_path,
        qa_pdf_path=qa_pdf_path,
        warnings=warnings,
        output_sharing=output_sharing,
    )


def build_wendy_wu_qbr_slides_payload(
    *,
    request_dir: str | Path,
    artifact: dict[str, Any],
    client_id: str = "wendy_wu",
    template_manifest: dict[str, Any] | None = None,
    project_root: str | Path = PROJECT_ROOT,
) -> dict[str, Any]:
    request_path = Path(request_dir)
    root = Path(project_root)
    config_loader = ConfigLoader(
        report_config_path=root / "config" / "report_config.yaml",
        chart_styles_path=root / "config" / "chart_styles.yaml",
        clients_config_path=root / "config" / "clients_config.json",
    )
    client_config = config_loader.get_client_config(client_id)
    source_manifest = _load_source_generation_manifest(request_path, artifact)
    performance_csv = _resolve_generated_file(
        source_manifest, request_path, "performance_csv"
    ) or request_path / "source_data" / "performance.csv"
    if not performance_csv.exists():
        raise FileNotFoundError(
            f"Could not find API-generated {client_config.get('name', client_id)} QBR performance.csv."
        )

    df = load_csv(performance_csv)
    quarter = detect_latest_complete_quarter(df)
    report = prepare_report_data(
        df,
        quarter,
        campaign_order=config_loader.get_campaign_types(client_config),
        destination_order=config_loader.get_destinations(client_config),
        destination_aliases=client_config.get("destination_aliases"),
        destination_other_config=client_config.get("destination_other"),
        report_mode="quarterly",
    )
    validate_report_data(report)

    chart_builder = ChartBuilder(
        request_path / "outputs" / "native_google_slides_qbr_charts",
        chart_styles=config_loader.get_chart_styles(client_config),
    )
    subtitle = _format_quarter_subtitle(quarter)
    footer = f"{quarter.label} | Summon Digital | Confidential"
    client_label = _client_label(client_id, client_config)
    currency = _currency_for_client(client_id, client_config)

    shape_text: dict[str, str] = {
        "p1_i20": client_label,
        "p1_i21": config_loader.get_report_title(client_config),
        "p1_i22": subtitle,
        "p1_i23": "Prepared by Summon",
        "p2_i45": f"{client_label} | {quarter.label} | Summon",
        "p6_i100": f"{client_label} | {quarter.label} | Summon",
        "p17_i402": f"{client_label} | {quarter.label} | Summon",
        "p28_i706": f"{client_label} | {quarter.label} | Summon",
        "p30_i731": f"{client_label} | {quarter.label} | Summon",
        "p33_i770": f"{client_label} | {quarter.label} | Summon",
        "p35_i795": f"{client_label} | {quarter.label} | Summon",
        "p38_i851": f"{client_label} | {quarter.label} PPC Report | Prepared by Summon",
        "p38_i852": footer,
    }
    tables: dict[str, dict[str, Any]] = {}
    charts: dict[tuple[str, str], Path] = {}
    warnings: list[str] = []

    _populate_cover(shape_text, report)
    _populate_summary_sections(
        shape_text=shape_text,
        tables=tables,
        charts=charts,
        report=report,
        subtitle=subtitle,
        footer=footer,
        chart_builder=chart_builder,
        currency_symbol=currency["symbol"],
    )
    _populate_monthly_chart_sections(
        charts=charts,
        report=report,
        chart_builder=chart_builder,
        currency_symbol=currency["symbol"],
    )
    _populate_trend_sections(
        shape_text=shape_text,
        charts=charts,
        tables=tables,
        request_path=request_path,
        artifact=artifact,
        source_manifest=source_manifest,
        client_config=client_config,
        config_loader=config_loader,
        quarter=quarter,
        subtitle=subtitle,
        chart_builder=chart_builder,
        warnings=warnings,
    )
    _populate_other_campaigns_section(
        charts=charts,
        request_path=request_path,
        source_manifest=source_manifest,
        client_config=client_config,
        chart_builder=chart_builder,
        warnings=warnings,
    )
    auction_sources = _resolve_manual_auction_sources(request_path)
    auction_path = _resolve_manual_auction_path(request_path)
    _populate_auction_section(
        shape_text=shape_text,
        tables=tables,
        auction_path=auction_path,
        auction_sources=auction_sources,
        client_config=client_config,
        config_loader=config_loader,
        subtitle=subtitle,
        warnings=warnings,
    )
    _populate_review_required_sections(shape_text, subtitle, footer, quarter.label)

    revenue_tiles = _revenue_tile_payload(report)
    _localize_payload_currency(shape_text, tables, currency["symbol"])
    for tile in revenue_tiles.values():
        tile["value"] = _localize_currency_text(tile["value"], currency["symbol"])

    manifest = template_manifest or _read_json(_template_manifest_path(client_id))
    return {
        "artifact": {
            **artifact,
            "period": {
                "label": quarter.label,
                "subtitle": subtitle,
                "date_range": {
                    "from": quarter.start.strftime("%Y-%m-%d"),
                    "to": quarter.end.strftime("%Y-%m-%d"),
                },
            },
        },
        "period": {
            "label": quarter.label,
            "subtitle": subtitle,
            "start": quarter.start.strftime("%Y-%m-%d"),
            "end": quarter.end.strftime("%Y-%m-%d"),
        },
        "shape_text": shape_text,
        "tables": tables,
        "revenue_tiles": revenue_tiles,
        "charts": charts,
        "performance_csv": str(performance_csv),
        "currency": currency,
        "manual_inputs": {
            "auction_insights_csv": str(auction_path) if auction_path else None,
            "google_ads_auction_insights_csv": str(auction_sources.get("Google Ads") or ""),
            "microsoft_ads_auction_insights_csv": str(auction_sources.get("Microsoft Ads") or ""),
            "auction_insights_required": True,
            "auction_insights_same_period_required": True,
            "other_campaign_exports_dir": str(_resolve_other_campaigns_dir(request_path, source_manifest) or ""),
            "other_campaign_exports_optional": True,
        },
        "automated_sources": {
            "ga4_performance_csv": str(performance_csv),
            "dataforseo_trends_ytd_current_dir": str(
                _resolve_generated_file(source_manifest, request_path, "trends_ytd_current_dir")
                or request_path / "source_data" / "trends_ytd_current"
            ),
            "dataforseo_trends_ytd_previous_dir": str(
                _resolve_generated_file(source_manifest, request_path, "trends_ytd_previous_dir")
                or request_path / "source_data" / "trends_ytd_previous"
            ),
        },
        "template_manifest": manifest,
        "warnings": warnings,
    }


def _effective_template_id(
    client_id: str,
    template: TemplateConfig,
    template_manifest: Mapping[str, Any],
    warnings: list[str],
) -> str:
    configured = str(template.template_id or "").strip()
    manifest_template = str(template_manifest.get("template_presentation_id") or "").strip()
    source = str(template_manifest.get("source_presentation_id") or "").strip()
    obsolete_ids = {source, *OBSOLETE_WENDY_WU_QBR_TEMPLATE_IDS.get(client_id, set())}
    if configured and configured in obsolete_ids and manifest_template:
        warnings.append(
            "Configured WWT QBR template points at an old/source deck; using the copied template deck instead."
        )
        return manifest_template
    return configured or manifest_template


def _template_manifest_path(client_id: str) -> Path:
    path = WENDY_WU_QBR_TEMPLATE_MANIFESTS.get(client_id)
    if not path:
        raise ValueError(f"No WWT QBR native template manifest configured for {client_id}.")
    return path


def _builder_name(client_id: str) -> str:
    return WENDY_WU_QBR_BUILDER_NAMES.get(client_id, "wendy_wu_qbr_template_manifest")


def _client_label(client_id: str, client_config: Mapping[str, Any]) -> str:
    if client_id == "wendy_wu":
        return "Wendy Wu Tours UK"
    return str(client_config.get("name") or "Wendy Wu Tours Australia")


def _currency_for_client(client_id: str, client_config: Mapping[str, Any]) -> dict[str, str]:
    country = str(client_config.get("country") or "").strip().lower()
    if client_id == "wendy_wu_australia" or country == "australia":
        return {"code": "AUD", "symbol": "$"}
    return {"code": "GBP", "symbol": "£"}


def _populate_cover(shape_text: dict[str, str], report: dict[str, Any]) -> None:
    kpis = _kpi_lookup(report["overall"])
    shape_text.update(
        {
            "p1_i25": _kpi_value(kpis, "Sales Leads"),
            "p1_i28": _kpi_value(kpis, "Cost"),
            "p1_i31": _kpi_value(kpis, "CPL"),
            "p1_i34": _kpi_value(kpis, "CVR"),
        }
    )


def _populate_summary_sections(
    *,
    shape_text: dict[str, str],
    tables: dict[str, dict[str, Any]],
    charts: dict[tuple[str, str], Path],
    report: dict[str, Any],
    subtitle: str,
    footer: str,
    chart_builder: ChartBuilder,
    currency_symbol: str,
) -> None:
    for key, section in SUMMARY_SLIDES.items():
        scope = _scope_for_section(report, section)
        if not scope:
            _populate_missing_summary(shape_text, section, subtitle, footer)
            continue

        shape_text[str(section["title_id"])] = _summary_title_for_section(key, section)
        shape_text[str(section["subtitle_id"])] = subtitle
        shape_text[str(section["footer_id"])] = footer
        kpis = _kpi_lookup(scope)
        for index, metric in enumerate(CARD_METRIC_ORDER):
            _maybe_set(shape_text, section["value_ids"], index, _kpi_value(kpis, metric))
            _maybe_set(
                shape_text,
                section["delta_ids"],
                index,
                f"YoY: {_kpi_yoy(kpis, metric)}",
            )
        bullets = (
            generate_overall_bullets(scope, report["mix_overall"])
            if key == "overall"
            else generate_scope_bullets(str(section["label"]), scope)
        )
        if section.get("bullets_id"):
            shape_text[str(section["bullets_id"])] = "\n".join(bullets)

        table_id = section.get("table_id")
        if table_id:
            scope_type = section["scope"][0]
            if scope_type == "destination":
                destination = str(section["scope"][1])
                table_df = _format_mix_table(report["dest_mix"].get(destination, pd.DataFrame()))
                tables[str(table_id)] = {
                    "values": _table_values(table_df, currency_symbol)
                }
            else:
                tables[str(table_id)] = {
                    "values": _table_values(
                        _format_monthly_table(scope["monthly"], currency_symbol)
                    )
                }

        chart_ids = section.get("chart_ids") or {}
        if "cpl_cvr" in chart_ids and "cost_leads" in chart_ids:
            scope_key = _chart_scope_key(key)
            charts[(str(chart_ids["cpl_cvr"]), "cpl_cvr")] = _build_cpl_cvr_chart(
                chart_builder, scope_key, scope["monthly"], currency_symbol
            )
            charts[(str(chart_ids["cost_leads"]), "cost_leads")] = _build_cost_leads_bar_chart(
                chart_builder, scope_key, scope["monthly"], currency_symbol
            )
        if "campaign_mix" in chart_ids:
            destination = str(section["scope"][1])
            charts[(str(chart_ids["campaign_mix"]), "campaign_mix")] = _build_combined_mix_chart(
                chart_builder,
                f"{_chart_scope_key(key)}_campaign_mix",
                report["dest_mix"].get(destination, pd.DataFrame()),
            )

    mix_charts = chart_builder.build_mix_charts("native_qbr_overall_mix", report["mix_overall"])
    charts[("p8_i151", "overall_cost_share")] = mix_charts["cost_share"]
    charts[("p8_i160", "overall_leads_share")] = mix_charts["leads_share"]
    shape_text["p8_i155"] = subtitle
    shape_text["p8_i157"] = footer
    shape_text["p8_i158"] = "\n".join(generate_mix_bullets(report["mix_overall"], "overall"))


def _populate_monthly_chart_sections(
    *,
    charts: dict[tuple[str, str], Path],
    report: dict[str, Any],
    chart_builder: ChartBuilder,
    currency_symbol: str,
) -> None:
    for key, section in MONTHLY_TREND_SLIDES.items():
        scope = _scope_for_section(report, section)
        if not scope:
            continue
        scope_key = _chart_scope_key(key)
        charts[(str(section["chart_ids"]["cpl_cvr"]), f"{key}_cpl_cvr")] = (
            _build_cpl_cvr_chart(chart_builder, scope_key, scope["monthly"], currency_symbol)
        )
        charts[(str(section["chart_ids"]["cost_leads"]), f"{key}_cost_leads")] = _build_cost_leads_bar_chart(
            chart_builder, scope_key, scope["monthly"], currency_symbol
        )


def _populate_trend_sections(
    *,
    shape_text: dict[str, str],
    charts: dict[tuple[str, str], Path],
    tables: dict[str, dict[str, Any]],
    request_path: Path,
    artifact: dict[str, Any],
    source_manifest: dict[str, Any] | None,
    client_config: dict[str, Any],
    config_loader: ConfigLoader,
    quarter: Any,
    subtitle: str,
    chart_builder: ChartBuilder,
    warnings: list[str],
) -> None:
    current_dir = _resolve_generated_file(source_manifest, request_path, "trends_ytd_current_dir")
    previous_dir = _resolve_generated_file(source_manifest, request_path, "trends_ytd_previous_dir")
    if current_dir is None:
        current_dir = request_path / "source_data" / "trends_ytd_current"
    if previous_dir is None:
        previous_dir = request_path / "source_data" / "trends_ytd_previous"

    current_df = TrendsLoader(current_dir).load_from_directory()
    previous_df = TrendsLoader(previous_dir).load_from_directory()
    if current_df.empty:
        warnings.append("DataForSEO YTD trend CSVs were not available for WWT QBR.")
        for section in TREND_SLIDES.values():
            shape_text[str(section["bullets_id"])] = (
                "Review required: DataForSEO trend source was not available."
            )
            shape_text[str(section["source_id"])] = "Source: DataForSEO Google Trends API"
        return

    trends_summary = summarize_trends(
        trends_df=current_df,
        quarter=quarter,
        brand_terms=client_config.get("brand_trends", {}).get("terms", []),
        destination_configs=client_config.get("destination_trends", {}).get("destinations", []),
        trend_aliases=client_config.get("trend_aliases", {}),
        comparison_period="ytd",
        previous_trends_df=previous_df if not previous_df.empty else None,
    )
    by_name = {
        str(summary.get("name")): summary
        for summary in trends_summary.get("destinations", [])
        if isinstance(summary, dict)
    }
    for key, section in TREND_SLIDES.items():
        if section["terms_kind"] == "brand":
            summary = trends_summary.get("brand")
        else:
            summary = by_name.get(str(section["destination"]))
        shape_text[str(section["source_id"])] = "Source: DataForSEO Google Trends API"
        if not summary:
            warnings.append(f"No YTD trend summary available for {section['label']}.")
            shape_text[str(section["bullets_id"])] = (
                f"Review required: no YTD trend source was returned for {section['label']}."
            )
            charts[(str(section["chart_id"]), f"{key}_trend")] = chart_builder._plot_empty_state(
                chart_builder.charts_dir / f"native_qbr_{key}_trend.png",
                f"No {section['label']} trend data",
            )
            continue
        charts[(str(section["chart_id"]), f"{key}_trend")] = chart_builder.build_trends_chart(
            f"native_qbr_{key}",
            summary["comparison"],
            str(section["chart_title"]),
            current_label=str(summary.get("current_series_label") or f"{quarter.year} YTD"),
            prior_label=str(summary.get("prior_series_label") or f"{quarter.year - 1} YTD"),
        )
        shape_text[str(section["bullets_id"])] = "\n".join(
            generate_trend_bullets(summary, str(section["label"]))
        )


def _populate_other_campaigns_section(
    *,
    charts: dict[tuple[str, str], Path],
    request_path: Path,
    source_manifest: dict[str, Any] | None,
    client_config: dict[str, Any],
    chart_builder: ChartBuilder,
    warnings: list[str],
) -> None:
    source_dir = _resolve_other_campaigns_dir(request_path, source_manifest)
    config = client_config.get("other_top_campaigns", {})
    if not config.get("enabled"):
        config = (
            get_wendy_wu_other_top_campaigns_config(str(client_config.get("id") or ""))
            or config
        )
    summary = load_other_campaign_summary(
        source_dir,
        exclude_terms=config.get("exclude_terms", []),
        top_n=int(config.get("top_n", 10)),
    )
    if not summary:
        warnings.append(
            "Other top-campaign exports were not available; generated empty review charts."
        )
        charts[("p27_i695", "other_top_clicks")] = chart_builder._plot_empty_state(
            chart_builder.charts_dir / "native_qbr_other_top_clicks.png",
            "No Other campaign data",
        )
        charts[("p27_i694", "other_top_conversions")] = chart_builder._plot_empty_state(
            chart_builder.charts_dir / "native_qbr_other_top_conversions.png",
            "No Other campaign data",
        )
        return

    built = chart_builder.build_other_top_campaign_charts(
        "native_qbr_other",
        summary["top_clicks"],
        summary["top_conversions"],
    )
    charts[("p27_i695", "other_top_clicks")] = built["top_clicks"]
    charts[("p27_i694", "other_top_conversions")] = built["top_conversions"]


def _populate_auction_section(
    *,
    shape_text: dict[str, str],
    tables: dict[str, dict[str, Any]],
    auction_path: Path | None,
    auction_sources: Mapping[str, Path],
    client_config: dict[str, Any],
    config_loader: ConfigLoader,
    subtitle: str,
    warnings: list[str],
) -> None:
    shape_text["p29_i714"] = AUCTION_TITLE
    shape_text["p29_i715"] = subtitle
    shape_text["p29_i718"] = _auction_source_note(AUCTION_PREFERRED_SOURCE)
    if not auction_sources and not auction_path:
        warnings.append(
            "Manual Google Ads Auction Insights CSV was not uploaded; auction slide marked review-required."
        )
        tables[AUCTION_TABLE_ID] = _auction_table_payload(
            [["Manual upload required", *["Review required"] * (len(AUCTION_HEADERS) - 1)]]
        )
        shape_text["p29_i719"] = (
            "Review required: export Non-Brand Auction Insights from Google Ads "
            "for the same QBR period and regenerate or update this slide."
        )
        return

    if auction_sources:
        if AUCTION_PREFERRED_SOURCE not in auction_sources:
            warnings.append(
                "Auction Insights is missing the manual Google Ads upload; "
                "the slide uses the available platform export instead."
            )
        auction_df = load_cross_platform_auction_csvs(auction_sources)
    else:
        auction_df = load_auction_csv(auction_path)
    auction_df, source_label = _select_auction_platform(auction_df)
    if source_label:
        shape_text["p29_i718"] = _auction_source_note(source_label)
    auction_df = _mark_own_auction_rows(
        auction_df, client_config.get("auction_insights", {}).get("client_domain")
    )
    summary = summarize_auction_insights(
        auction_df,
        client_domain=AUCTION_OWN_DOMAIN,
        known_competitors=client_config.get("auction_insights", {}).get(
            "known_competitors", []
        ),
    )
    if not summary:
        warnings.append(
            "Manual Auction Insights CSV was uploaded but no usable summary could be generated."
        )
        tables[AUCTION_TABLE_ID] = _auction_table_payload(
            [["No usable rows", *["n/a"] * (len(AUCTION_HEADERS) - 1)]]
        )
        shape_text["p29_i719"] = "Review required: Auction Insights CSV had no usable rows."
        return

    tables[AUCTION_TABLE_ID] = _auction_table_payload(_auction_table_rows(auction_df))
    shape_text["p29_i719"] = "\n".join(generate_auction_bullets(summary))


def _auction_source_note(source_label: str) -> str:
    return f"Source: {source_label} Auction Insights"


def _select_auction_platform(auction_df: pd.DataFrame) -> tuple[pd.DataFrame, str | None]:
    """Keep one platform so the slide table needs no Source column.

    Google Ads is preferred; another platform is only used when Google Ads rows are absent.
    """
    if "source" not in auction_df.columns or auction_df.empty:
        return auction_df, None
    sources = [str(value) for value in auction_df["source"].dropna().unique()]
    if not sources:
        return auction_df, None
    selected = AUCTION_PREFERRED_SOURCE if AUCTION_PREFERRED_SOURCE in sources else sources[0]
    platform_df = auction_df[auction_df["source"] == selected].drop(columns=["source"])
    return platform_df.reset_index(drop=True), selected


def _mark_own_auction_rows(auction_df: pd.DataFrame, client_domain: Any) -> pd.DataFrame:
    if auction_df.empty or "domain" not in auction_df.columns:
        return auction_df
    own_domains = {AUCTION_OWN_DOMAIN}
    client_key = str(client_domain or "").strip().lower().replace("www.", "").strip("/")
    if client_key:
        own_domains.add(client_key)
    marked = auction_df.copy()
    marked["domain"] = marked["domain"].map(
        lambda value: AUCTION_OWN_DOMAIN if str(value).strip().lower() in own_domains else value
    )
    return marked


def _auction_table_rows(auction_df: pd.DataFrame) -> list[list[str]]:
    """Own row first, then competitors by impression share (highest first)."""
    own_df = auction_df[auction_df["domain"] == AUCTION_OWN_DOMAIN].head(1)
    competitor_df = auction_df[auction_df["domain"] != AUCTION_OWN_DOMAIN]
    if "impression_share" in competitor_df.columns:
        competitor_df = competitor_df.sort_values(
            "impression_share", ascending=False, na_position="last", kind="stable"
        )
    competitor_limit = AUCTION_MAX_BODY_ROWS - len(own_df)
    ordered = pd.concat([own_df, competitor_df.head(competitor_limit)], ignore_index=True)
    rows = []
    for _, row in ordered.iterrows():
        domain = str(row["domain"])
        rows.append(
            [
                "You" if domain == AUCTION_OWN_DOMAIN else domain,
                *(_format_auction_pct(row.get(column)) for column in AUCTION_VALUE_COLUMNS),
            ]
        )
    return rows


def _format_auction_pct(value: Any) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{float(value) * 100:.1f}%"


def _auction_table_payload(body_rows: Sequence[Sequence[str]]) -> dict[str, Any]:
    return {
        "values": [list(AUCTION_HEADERS), *(list(row) for row in body_rows)],
        "header_fill_rgb": BLACK_RGB,
        # Reset template row fills so only the client row is highlighted.
        "body_fill_rgb": WHITE_RGB,
        "border_rgb": AUCTION_BORDER_RGB,
        "highlight_rows": [
            index
            for index, row in enumerate(body_rows, start=1)
            if row and str(row[0]).strip().lower() == AUCTION_OWN_DOMAIN
        ],
        "highlight_fill_rgb": AUCTION_HIGHLIGHT_RGB,
    }


def _table_style_requests(
    table_id: str, table_payload: Mapping[str, Any]
) -> list[dict[str, Any]]:
    values = table_payload.get("values") or []
    column_count = max((len(row) for row in values), default=0)
    if column_count <= 0:
        return []
    fills: list[tuple[int, dict[str, float]]] = []
    if table_payload.get("header_fill_rgb"):
        fills.append((0, table_payload["header_fill_rgb"]))
    if table_payload.get("body_fill_rgb"):
        fills.extend(
            (row_index, table_payload["body_fill_rgb"])
            for row_index in range(1, len(values))
        )
    highlight_rgb = table_payload.get("highlight_fill_rgb")
    if highlight_rgb:
        fills.extend(
            (int(row_index), highlight_rgb)
            for row_index in table_payload.get("highlight_rows") or []
            if 0 < int(row_index) < len(values)
        )
    requests_body = [
        {
            "updateTableCellProperties": {
                "objectId": table_id,
                "tableRange": {
                    "location": {"rowIndex": row_index, "columnIndex": 0},
                    "rowSpan": 1,
                    "columnSpan": column_count,
                },
                "tableCellProperties": {
                    "tableCellBackgroundFill": {"solidFill": {"color": {"rgbColor": rgb}}}
                },
                "fields": "tableCellBackgroundFill.solidFill.color",
            }
        }
        for row_index, rgb in fills
    ]
    border_rgb = table_payload.get("border_rgb")
    if border_rgb:
        # Overwrite leftover template border colours (e.g. old highlight rows).
        requests_body.append(
            {
                "updateTableBorderProperties": {
                    "objectId": table_id,
                    "tableRange": {
                        "location": {"rowIndex": 0, "columnIndex": 0},
                        "rowSpan": len(values),
                        "columnSpan": column_count,
                    },
                    "borderPosition": "ALL",
                    "tableBorderProperties": {
                        "tableBorderFill": {"solidFill": {"color": {"rgbColor": border_rgb}}},
                        "weight": {"magnitude": AUCTION_BORDER_WEIGHT_EMU, "unit": "EMU"},
                        "dashStyle": "SOLID",
                    },
                    "fields": "tableBorderFill.solidFill.color,weight,dashStyle",
                }
            }
        )
    return requests_body


def _populate_review_required_sections(
    shape_text: dict[str, str], subtitle: str, footer: str, period_label: str
) -> None:
    shape_text.update(
        {
            "p31_i740": subtitle,
            "p31_i742": footer,
            "p31_i743": "Review required: Testing section to be completed by the account team.",
            "p32_i755": subtitle,
            "p32_i757": footer,
            "p32_i758": "Review required: Testing section to be completed by the account team.",
            "p34_i782": footer,
            "p34_i783": "Review required: Other Updates to be completed by the account team.",
            "p36_i803": f"Next Steps for {period_label}",
            "p36_i804": subtitle,
            "p36_i806": footer,
            "p37_i831": f"Next Steps for {period_label}",
            "p37_i832": subtitle,
            "p37_i834": footer,
            "p37_i838": "Review required: next-step continuation placeholder to be completed by the account team.",
        }
    )


def _scope_for_section(report: dict[str, Any], section: Mapping[str, Any]) -> dict[str, Any] | None:
    scope_type, name = section["scope"]
    if scope_type == "overall":
        return report.get("overall")
    if scope_type == "campaign":
        return report.get("campaigns", {}).get(name)
    if scope_type == "destination":
        return report.get("destinations", {}).get(name)
    return None


def _populate_missing_summary(
    shape_text: dict[str, str],
    section: Mapping[str, Any],
    subtitle: str,
    footer: str,
) -> None:
    shape_text[str(section["title_id"])] = _summary_title_for_section("", section)
    shape_text[str(section["subtitle_id"])] = subtitle
    shape_text[str(section["footer_id"])] = footer
    for object_id in section.get("value_ids", []):
        shape_text[str(object_id)] = "n/a"
    for object_id in section.get("delta_ids", []):
        shape_text[str(object_id)] = "YoY: n/a"
    if section.get("bullets_id"):
        shape_text[str(section["bullets_id"])] = (
            "No data is available for this section in the generated performance CSV."
        )


def _summary_title_for_section(key: str, section: Mapping[str, Any]) -> str:
    label = str(section.get("label") or "").strip()
    scope_type = str((section.get("scope") or ("", None))[0])
    if key == "overall":
        return "Overall Performance Trend"
    if scope_type == "campaign":
        return f"{label} Summary"
    if scope_type == "destination":
        prefix = "Other (Destination)" if label == "Other" else label
        return f"{prefix} Summary + YoY"
    return label


def _format_monthly_table(monthly_df: pd.DataFrame, currency_symbol: str = "£") -> pd.DataFrame:
    if not isinstance(monthly_df, pd.DataFrame) or monthly_df.empty:
        return pd.DataFrame(
            columns=[
                "Month",
                "Impressions",
                "Clicks",
                "CTR",
                "CPC",
                "Cost",
                "Sales Leads",
                "CPL",
                "CVR",
            ]
        )
    columns = [
        "Month",
        "Impressions",
        "Clicks",
        "CTR",
        "CPC",
        "Cost",
        "Sales Leads",
        "CPL",
        "CVR",
    ]
    working = monthly_df[[column for column in columns if column in monthly_df.columns]].copy()
    for col in ["Impressions", "Clicks", "Sales Leads"]:
        if col in working.columns:
            working[col] = working[col].map(lambda value: f"{int(round(float(value))):,}")
    for col in ["Cost", "CPC", "CPL"]:
        if col in working.columns:
            working[col] = working[col].map(
                lambda value: _fmt_currency(value, currency_symbol)
            )
    for col in ["CTR", "CVR"]:
        if col in working.columns:
            working[col] = working[col].map(_fmt_percent)
    return working[columns]


def _table_values(table_df: pd.DataFrame, currency_symbol: str = "£") -> list[list[str]]:
    if table_df.empty:
        return [["Status"], ["No data available"]]
    values = [list(map(str, table_df.columns))]
    for row in table_df.fillna("").astype(str).itertuples(index=False):
        values.append(
            [_localize_currency_text(str(value), currency_symbol) for value in row]
        )
    return values


def _build_scalar_text_requests(
    shape_text: Mapping[str, str], presentation: dict[str, Any]
) -> list[dict[str, Any]]:
    existing_ids = _text_object_ids(presentation)
    requests_body: list[dict[str, Any]] = []
    for object_id, text in shape_text.items():
        if object_id not in existing_ids:
            continue
        requests_body.append(
            {"deleteText": {"objectId": object_id, "textRange": {"type": "ALL"}}}
        )
        if str(text):
            requests_body.append(
                {
                    "insertText": {
                        "objectId": object_id,
                        "insertionIndex": 0,
                        "text": str(text),
                    }
                }
            )
    return requests_body


def _replace_existing_table_exact_requests(
    *,
    table_id: str,
    values: Sequence[Sequence[str]],
    existing_rows: int,
    existing_columns: int,
    existing_cell_text: dict[tuple[int, int], str],
    column_widths: Sequence[int] | None = None,
) -> list[dict[str, Any]]:
    if existing_rows <= 0 or existing_columns <= 0:
        return []
    target_rows = max(1, len(values))
    target_columns = max(1, max(len(row) for row in values) if values else 1)
    requests_body: list[dict[str, Any]] = []

    if target_rows > existing_rows:
        requests_body.append(
            {
                "insertTableRows": {
                    "tableObjectId": table_id,
                    "cellLocation": {"rowIndex": existing_rows - 1, "columnIndex": 0},
                    "insertBelow": True,
                    "number": target_rows - existing_rows,
                }
            }
        )
    if target_columns > existing_columns:
        requests_body.append(
            {
                "insertTableColumns": {
                    "tableObjectId": table_id,
                    "cellLocation": {
                        "rowIndex": 0,
                        "columnIndex": existing_columns - 1,
                    },
                    "insertRight": True,
                    "number": target_columns - existing_columns,
                }
            }
        )
    if target_columns < existing_columns:
        for _ in range(existing_columns - target_columns):
            requests_body.append(
                {
                    "deleteTableColumn": {
                        "tableObjectId": table_id,
                        "cellLocation": {
                            "rowIndex": 0,
                            "columnIndex": target_columns,
                        },
                    }
                }
            )
    if target_rows < existing_rows:
        for _ in range(existing_rows - target_rows):
            requests_body.append(
                {
                    "deleteTableRow": {
                        "tableObjectId": table_id,
                        "cellLocation": {
                            "rowIndex": target_rows,
                            "columnIndex": 0,
                        },
                    }
                }
            )

    requests_body.extend(
        _table_cell_text_requests(
            table_id, values, target_rows, target_columns, existing_cell_text
        )
    )
    requests_body.extend(
        _table_column_width_requests(
            table_id,
            _target_column_widths(target_columns, column_widths or []),
        )
    )
    requests_body.extend(_table_format_requests(table_id, values, target_columns))
    return requests_body


def _target_column_widths(column_count: int, existing_widths: Sequence[int]) -> list[int]:
    widths = [int(width) for width in existing_widths if int(width) > 0]
    if not widths:
        return []
    total = sum(widths)
    if total <= 0:
        return widths[:column_count]
    weights = {
        6: (0.22, 0.19, 0.14, 0.15, 0.15, 0.15),
        7: (0.22, 0.13, 0.13, 0.15, 0.13, 0.12, 0.12),
        8: (0.11, 0.19, 0.12, 0.12, 0.15, 0.11, 0.10, 0.10),
        9: (0.12, 0.13, 0.10, 0.09, 0.09, 0.13, 0.10, 0.10, 0.14),
    }.get(column_count)
    if not weights:
        return widths[:column_count]
    adjusted = [int(round(total * weight)) for weight in weights]
    adjusted[-1] += total - sum(adjusted)
    return adjusted


def _table_widths_by_id(presentation: dict[str, Any]) -> dict[str, list[int]]:
    widths: dict[str, list[int]] = {}
    for slide in presentation.get("slides") or []:
        for element in slide.get("pageElements") or []:
            table = element.get("table")
            object_id = str(element.get("objectId") or "")
            if not object_id or not isinstance(table, dict):
                continue
            values = []
            for column in table.get("tableColumns") or []:
                width = (
                    (column.get("columnWidth") or {}).get("magnitude")
                    if isinstance(column, dict)
                    else None
                )
                if isinstance(width, (int, float)):
                    values.append(int(width))
            widths[object_id] = values
    return widths


def _upload_chart_assets(
    asset_store: DriveChartAssetStore, charts: Mapping[tuple[str, str], Path]
) -> dict[str, str]:
    urls: dict[str, str] = {}
    for (image_id, _chart_key), path in charts.items():
        if not Path(path).exists():
            continue
        asset = asset_store.upload_chart(path)
        urls[str(image_id)] = asset.public_url
    return urls


def _build_chart_requests(uploaded_assets: Mapping[str, str]) -> list[dict[str, Any]]:
    requests_body = []
    for image_id, url in uploaded_assets.items():
        requests_body.append(
            {
                "replaceImage": {
                    "imageObjectId": image_id,
                    "url": url,
                    "imageReplaceMethod": "CENTER_INSIDE",
                }
            }
        )
    return requests_body


def _build_template_white_text_style_requests(
    presentation: Mapping[str, Any], template_manifest: Mapping[str, Any]
) -> list[dict[str, Any]]:
    page_height = _page_height_magnitude(presentation)
    header_limit = page_height * 0.095
    footer_start = page_height * 0.94
    force_white_slide_ids = _full_black_slide_ids(template_manifest)
    requests_body: list[dict[str, Any]] = []
    styled_ids: set[str] = set()

    for slide in presentation.get("slides") or []:
        slide_id = str(slide.get("objectId") or "")
        force_slide = slide_id in force_white_slide_ids
        for element in slide.get("pageElements") or []:
            object_id = str(element.get("objectId") or "")
            if (
                not object_id
                or object_id in styled_ids
                or not _element_has_shape_text(element)
            ):
                continue
            y_position = _element_translate_y(element)
            if force_slide or y_position <= header_limit or y_position >= footer_start:
                requests_body.append(_text_color_request(object_id, WHITE_RGB))
                styled_ids.add(object_id)
    return requests_body


def _full_black_slide_ids(template_manifest: Mapping[str, Any]) -> set[str]:
    slide_ids: set[str] = set()
    for section in template_manifest.get("sections") or []:
        if not isinstance(section, Mapping):
            continue
        role = str(section.get("role") or "")
        slide_id = str(section.get("slide_id") or "")
        if slide_id and (role in {"cover", "closing"} or role.endswith("_divider")):
            slide_ids.add(slide_id)
    return slide_ids


def _element_has_shape_text(element: Mapping[str, Any]) -> bool:
    shape = element.get("shape")
    if not isinstance(shape, Mapping):
        return False
    text = shape.get("text")
    return isinstance(text, Mapping) and bool(text.get("textElements"))


def _element_translate_y(element: Mapping[str, Any]) -> float:
    transform = element.get("transform")
    if not isinstance(transform, Mapping):
        return 0.0
    value = transform.get("translateY", 0)
    return float(value) if isinstance(value, (int, float)) else 0.0


def _page_height_magnitude(presentation: Mapping[str, Any]) -> float:
    page_size = presentation.get("pageSize")
    if isinstance(page_size, Mapping):
        height = page_size.get("height")
        if isinstance(height, Mapping):
            magnitude = height.get("magnitude")
            if isinstance(magnitude, (int, float)) and magnitude > 0:
                return float(magnitude)
    return 5_143_500.0


def _build_cost_delta_style_requests(presentation: Mapping[str, Any]) -> list[dict[str, Any]]:
    existing_ids = _text_object_ids(dict(presentation))
    requests_body = []
    for section in SUMMARY_SLIDES.values():
        delta_ids = list(section.get("delta_ids") or [])
        if len(delta_ids) < 2:
            continue
        delta_id = str(delta_ids[1])
        if delta_id in existing_ids:
            requests_body.append(_text_color_request(delta_id, NEUTRAL_DELTA_RGB))
    return requests_body


def _revenue_tile_payload(report: Mapping[str, Any]) -> dict[str, dict[str, str]]:
    """Value and YoY text for the Revenue tile on each summary slide."""
    if not report.get("include_revenue"):
        return {}
    tiles: dict[str, dict[str, str]] = {}
    for key, section in SUMMARY_SLIDES.items():
        scope = _scope_for_section(report, section)
        # Quarterly scope KPIs omit Revenue, but totals and YoY carry it.
        kpis = (
            _kpi_lookup({"kpis": build_kpi_summary(scope["total"], scope["yoy"], include_revenue=True)})
            if scope
            else {}
        )
        tiles[key] = {
            "value": _whole_currency(_kpi_value(kpis, REVENUE_TILE_METRIC)),
            "yoy": f"YoY: {_kpi_yoy(kpis, REVENUE_TILE_METRIC)}",
        }
    return tiles


def _whole_currency(value: str) -> str:
    # Quarterly revenue runs to seven figures; pence would not fit a narrowed tile.
    return re.sub(r"(\d)\.\d{2}$", r"\1", str(value))


def _revenue_tile_requests(
    presentation: Mapping[str, Any], revenue_tiles: Mapping[str, Mapping[str, str]]
) -> list[dict[str, Any]]:
    """Narrow each summary slide's six KPI tiles and append a seventh Revenue tile."""
    elements = _top_level_elements_by_id(presentation)
    requests_body: list[dict[str, Any]] = []
    for key, tile_text in revenue_tiles.items():
        section = SUMMARY_SLIDES.get(key)
        if not section:
            continue
        tiles = _summary_tiles(elements, section)
        if tiles is None:
            continue
        requests_body.extend(_seven_tile_requests(tiles, section, tile_text, elements))
    return requests_body


def _top_level_elements_by_id(presentation: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    elements: dict[str, dict[str, Any]] = {}
    for slide in presentation.get("slides") or []:
        slide_id = str(slide.get("objectId") or "")
        for element in slide.get("pageElements") or []:
            box = _element_box(element)
            if box is None or "shape" not in element:
                continue
            elements[str(element.get("objectId"))] = {
                "element": element,
                "slide_id": slide_id,
                "box": box,
                "has_text": bool(_shape_plain_text(element).strip()),
            }
    return elements


def _element_box(element: Mapping[str, Any]) -> tuple[float, float, float, float] | None:
    transform = element.get("transform") or {}
    size = element.get("size") or {}
    width = (size.get("width") or {}).get("magnitude")
    height = (size.get("height") or {}).get("magnitude")
    if not width or not height or transform.get("shearX") or transform.get("shearY"):
        return None
    return (
        float(transform.get("translateX", 0.0)),
        float(transform.get("translateY", 0.0)),
        float(width) * float(transform.get("scaleX", 1.0)),
        float(height) * float(transform.get("scaleY", 1.0)),
    )


def _shape_plain_text(element: Mapping[str, Any]) -> str:
    text = ((element.get("shape") or {}).get("text") or {}).get("textElements") or []
    return "".join((item.get("textRun") or {}).get("content", "") for item in text)


def _summary_tiles(
    elements: Mapping[str, Mapping[str, Any]], section: Mapping[str, Any]
) -> list[list[str]] | None:
    """Group the slide's shapes into the six tiles: each card plus everything drawn on it."""
    value_ids = [str(value_id) for value_id in section.get("value_ids") or []]
    if len(value_ids) != len(CARD_METRIC_ORDER) or any(v not in elements for v in value_ids):
        return None
    slide_id = elements[value_ids[0]]["slide_id"]
    slide_elements = {k: v for k, v in elements.items() if v["slide_id"] == slide_id}
    tiles: list[list[str]] = []
    for value_id in value_ids:
        vx, vy, vw, vh = slide_elements[value_id]["box"]
        centre = (vx + vw / 2, vy + vh / 2)
        cards = [
            (item["box"][2] * item["box"][3], object_id)
            for object_id, item in slide_elements.items()
            if not item["has_text"] and _box_contains(item["box"], centre)
        ]
        if not cards:
            return None
        card_id = max(cards)[1]
        card_box = slide_elements[card_id]["box"]
        members = [
            object_id
            for object_id, item in slide_elements.items()
            if _box_contains(card_box, _box_centre(item["box"]))
        ]
        tiles.append([card_id] + [m for m in members if m != card_id])
    return tiles


def _box_centre(box: tuple[float, float, float, float]) -> tuple[float, float]:
    return box[0] + box[2] / 2, box[1] + box[3] / 2


def _box_contains(box: tuple[float, float, float, float], point: tuple[float, float]) -> bool:
    x, y, w, h = box
    return x <= point[0] <= x + w and y <= point[1] <= y + h


def _seven_tile_requests(
    tiles: list[list[str]],
    section: Mapping[str, Any],
    tile_text: Mapping[str, str],
    elements: Mapping[str, Mapping[str, Any]],
) -> list[dict[str, Any]]:
    value_ids = [str(value_id) for value_id in section["value_ids"]]
    delta_ids = [str(delta_id) for delta_id in section["delta_ids"]]
    card_boxes = [elements[tile[0]]["box"] for tile in tiles]
    left = card_boxes[0][0]
    right = card_boxes[-1][0] + card_boxes[-1][2]
    gap = max(0.0, card_boxes[1][0] - (card_boxes[0][0] + card_boxes[0][2]))
    tile_count = len(tiles) + 1
    tile_width = (right - left - gap * (tile_count - 1)) / tile_count

    source_tile = tiles[-1]
    new_ids = {object_id: f"{object_id}{REVENUE_TILE_ID_SUFFIX}" for object_id in source_tile}
    requests_body: list[dict[str, Any]] = [
        {"duplicateObject": {"objectId": object_id, "objectIds": {object_id: new_id}}}
        for object_id, new_id in new_ids.items()
    ]

    placements = [(tile, value_ids[index], {}) for index, tile in enumerate(tiles)]
    placements.append((source_tile, value_ids[-1], new_ids))
    for index, (tile, value_id, rename) in enumerate(placements):
        card_x, _card_y, card_w, _card_h = elements[tile[0]]["box"]
        scale = tile_width / card_w
        tile_left = left + index * (tile_width + gap)
        for object_id in tile:
            x, y, w, _h = elements[object_id]["box"]
            if object_id == value_id:
                new_width = tile_width * TILE_VALUE_WIDTH_SHARE
                new_x = tile_left + (tile_width - new_width) / 2
            else:
                new_width = w * scale
                new_x = tile_left + (x - card_x) * scale
            requests_body.append(
                _absolute_transform_request(
                    rename.get(object_id, object_id), elements[object_id]["element"], new_x, y, new_width
                )
            )

    source_value, source_delta = value_ids[-1], delta_ids[-1]
    label_ids = [
        object_id
        for object_id in source_tile
        if elements[object_id]["has_text"] and object_id not in {source_value, source_delta}
    ]
    texts = {new_ids[source_value]: tile_text["value"]}
    if source_delta in new_ids:
        texts[new_ids[source_delta]] = tile_text["yoy"]
    if label_ids:
        texts[new_ids[label_ids[0]]] = REVENUE_TILE_METRIC
    for object_id, text in texts.items():
        requests_body.append({"deleteText": {"objectId": object_id, "textRange": {"type": "ALL"}}})
        requests_body.append(
            {"insertText": {"objectId": object_id, "insertionIndex": 0, "text": str(text)}}
        )
    if label_ids:
        # Values and deltas on the other tiles are rewritten too and fall back to the
        # default style; labels keep the template styling, so restore it on ours.
        style_request = _copy_text_style_request(
            new_ids[label_ids[0]], elements[label_ids[0]]["element"]
        )
        if style_request:
            requests_body.append(style_request)
    if source_delta in new_ids:
        requests_body.append(
            _text_color_request(new_ids[source_delta], _delta_rgb(tile_text["yoy"]))
        )

    for value_id in [*value_ids, new_ids[source_value]]:
        requests_body.append(
            {
                "updateTextStyle": {
                    "objectId": value_id,
                    "textRange": {"type": "ALL"},
                    "style": {"fontSize": {"magnitude": TILE_VALUE_FONT_PT, "unit": "PT"}},
                    "fields": "fontSize",
                }
            }
        )

    card_height = elements[source_tile[0]]["box"][3]
    for object_id in source_tile[1:]:
        if not elements[object_id]["has_text"] and elements[object_id]["box"][3] < card_height * 0.2:
            requests_body.append(
                {
                    "updateShapeProperties": {
                        "objectId": new_ids[object_id],
                        "shapeProperties": {
                            "shapeBackgroundFill": {
                                "solidFill": {"color": {"rgbColor": dict(REVENUE_TILE_ACCENT_RGB)}}
                            },
                            "outline": {
                                "outlineFill": {
                                    "solidFill": {"color": {"rgbColor": dict(REVENUE_TILE_ACCENT_RGB)}}
                                }
                            },
                        },
                        "fields": "shapeBackgroundFill.solidFill.color,outline.outlineFill.solidFill.color",
                    }
                }
            )
    return requests_body


def _copy_text_style_request(object_id: str, source: Mapping[str, Any]) -> dict[str, Any] | None:
    text_elements = ((source.get("shape") or {}).get("text") or {}).get("textElements") or []
    style = next(
        (
            (item.get("textRun") or {}).get("style")
            for item in text_elements
            if (item.get("textRun") or {}).get("content", "").strip()
        ),
        None,
    )
    if not style:
        return None
    copied = {
        field: style[field]
        for field in ("foregroundColor", "fontSize", "fontFamily", "bold")
        if field in style
    }
    if not copied:
        return None
    return {
        "updateTextStyle": {
            "objectId": object_id,
            "textRange": {"type": "ALL"},
            "style": copied,
            "fields": ",".join(copied),
        }
    }


def _absolute_transform_request(
    object_id: str, element: Mapping[str, Any], x: float, y: float, width: float
) -> dict[str, Any]:
    transform = element.get("transform") or {}
    base_width = float(((element.get("size") or {}).get("width") or {}).get("magnitude") or 1.0)
    return {
        "updatePageElementTransform": {
            "objectId": object_id,
            "applyMode": "ABSOLUTE",
            "transform": {
                "scaleX": width / base_width,
                "scaleY": float(transform.get("scaleY", 1.0)),
                "shearX": 0,
                "shearY": 0,
                "translateX": round(x),
                "translateY": round(y),
                "unit": "EMU",
            },
        }
    }


def _delta_rgb(label: str) -> dict[str, float]:
    match = re.search(r"([+-])\d", str(label))
    if not match:
        return dict(NEUTRAL_DELTA_RGB)
    return dict(POSITIVE_RGB if match.group(1) == "+" else NEGATIVE_RGB)


def _text_color_request(object_id: str, rgb: Mapping[str, float]) -> dict[str, Any]:
    return {
        "updateTextStyle": {
            "objectId": object_id,
            "textRange": {"type": "ALL"},
            "style": {
                "foregroundColor": {"opaqueColor": {"rgbColor": dict(rgb)}},
            },
            "fields": "foregroundColor",
        }
    }


def _build_cpl_cvr_chart(
    chart_builder: ChartBuilder,
    scope_key: str,
    monthly_table: pd.DataFrame,
    currency_symbol: str,
) -> Path:
    out_path = chart_builder.charts_dir / f"{scope_key}_cpl_cvr.png"
    if (
        not isinstance(monthly_table, pd.DataFrame)
        or monthly_table.empty
        or "Month" not in monthly_table.columns
        or "CPL" not in monthly_table.columns
        or "CVR" not in monthly_table.columns
    ):
        return chart_builder._plot_empty_state(out_path, "No CPL or CVR data")

    df = monthly_table[monthly_table["Month"] != "Total"].copy()
    if df.empty:
        return chart_builder._plot_empty_state(out_path, "No CPL or CVR data")

    months = df["Month"].astype(str).tolist()
    cpl = pd.to_numeric(df["CPL"], errors="coerce").tolist()
    cvr = [
        value * 100 if value is not None and not pd.isna(value) else None
        for value in pd.to_numeric(df["CVR"], errors="coerce").tolist()
    ]

    fig, ax_cpl = plt.subplots(figsize=chart_builder.figure_size)
    ax_cvr = ax_cpl.twinx()

    cpl_line = ax_cpl.plot(
        months,
        cpl,
        marker="o",
        markersize=7,
        color=chart_builder.colors.get("cpl", "#C32026"),
        linewidth=2.5,
        label=f"CPL ({currency_symbol})",
        zorder=3,
    )[0]
    cvr_line = ax_cvr.plot(
        months,
        cvr,
        marker="o",
        markersize=7,
        color=chart_builder.colors.get("cvr", "#111111"),
        linewidth=3,
        label="CVR (%)",
        zorder=4,
    )[0]

    ax_cpl.set_title("CPL vs CVR", fontsize=chart_builder.title_size)
    ax_cpl.set_xlabel("Month", fontsize=chart_builder.body_size)
    ax_cpl.set_ylabel(f"CPL ({currency_symbol})", fontsize=chart_builder.body_size)
    ax_cvr.set_ylabel("CVR (%)", fontsize=chart_builder.body_size)
    ax_cpl.tick_params(axis="both", labelsize=chart_builder.body_size)
    ax_cvr.tick_params(axis="both", labelsize=chart_builder.body_size)
    ax_cpl.grid(axis="y", alpha=0.2)

    for index, value in enumerate(cpl):
        if value is None or pd.isna(value):
            continue
        ax_cpl.annotate(
            _format_currency_label(float(value), currency_symbol),
            xy=(index, value),
            xytext=(0, 10),
            textcoords="offset points",
            ha="center",
            fontsize=chart_builder.body_size - 1,
            color=chart_builder.colors.get("cpl", "#C32026"),
        )

    for index, value in enumerate(cvr):
        if value is None or pd.isna(value):
            continue
        ax_cvr.annotate(
            f"{value:.1f}%",
            xy=(index, value),
            xytext=(0, -16),
            textcoords="offset points",
            ha="center",
            fontsize=chart_builder.body_size - 1,
            color=chart_builder.colors.get("cvr", "#111111"),
        )

    ax_cpl.legend(
        [cpl_line, cvr_line],
        [f"CPL ({currency_symbol})", "CVR (%)"],
        loc="upper left",
        fontsize=chart_builder.body_size,
    )
    plt.tight_layout()
    fig.savefig(out_path, dpi=180)
    plt.close(fig)
    return out_path


def _build_cost_leads_bar_chart(
    chart_builder: ChartBuilder,
    scope_key: str,
    monthly_table: pd.DataFrame,
    currency_symbol: str,
) -> Path:
    out_path = chart_builder.charts_dir / f"{scope_key}_cost_leads_bars.png"
    if (
        not isinstance(monthly_table, pd.DataFrame)
        or monthly_table.empty
        or "Month" not in monthly_table.columns
        or "Cost" not in monthly_table.columns
        or "Sales Leads" not in monthly_table.columns
    ):
        return chart_builder._plot_empty_state(out_path, "No cost or lead data")

    df = monthly_table[monthly_table["Month"] != "Total"].copy()
    if df.empty:
        return chart_builder._plot_empty_state(out_path, "No cost or lead data")

    months = df["Month"].astype(str).tolist()
    cost = pd.to_numeric(df["Cost"], errors="coerce").fillna(0).tolist()
    leads = pd.to_numeric(df["Sales Leads"], errors="coerce").fillna(0).tolist()
    x_positions = list(range(len(months)))
    width = 0.34

    fig, ax_leads = plt.subplots(figsize=chart_builder.figure_size)
    ax_cost = ax_leads.twinx()

    lead_bars = ax_leads.bar(
        [position - width / 2 for position in x_positions],
        leads,
        width=width,
        color=chart_builder.colors.get("leads", "#111111"),
        alpha=0.92,
        label="Sales Leads",
        zorder=3,
    )
    cost_bars = ax_cost.bar(
        [position + width / 2 for position in x_positions],
        cost,
        width=width,
        color=chart_builder.colors.get("cost", "#D83A40"),
        alpha=0.9,
        label=f"Cost ({currency_symbol})",
        zorder=2,
    )

    ax_leads.set_title("Cost vs Sales Leads", fontsize=chart_builder.title_size)
    ax_leads.set_xlabel("Month", fontsize=chart_builder.body_size)
    ax_leads.set_ylabel("Sales Leads", fontsize=chart_builder.body_size)
    ax_cost.set_ylabel(f"Cost ({currency_symbol})", fontsize=chart_builder.body_size)
    ax_leads.set_xticks(x_positions, months)
    ax_leads.tick_params(axis="both", labelsize=chart_builder.body_size)
    ax_cost.tick_params(axis="y", labelsize=chart_builder.body_size)
    ax_leads.grid(axis="y", alpha=0.2)
    ax_leads.set_axisbelow(True)
    ax_cost.yaxis.set_major_formatter(
        FuncFormatter(lambda value, _: _format_currency_axis(value, currency_symbol))
    )
    ax_leads.yaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value:,.0f}"))

    max_leads = max([float(value) for value in leads] + [0.0])
    max_cost = max([float(value) for value in cost] + [0.0])
    ax_leads.set_ylim(top=max_leads * 1.18 if max_leads > 0 else 1)
    ax_cost.set_ylim(top=max_cost * 1.18 if max_cost > 0 else 1)

    for bar, value in zip(lead_bars, leads):
        if not value:
            continue
        ax_leads.annotate(
            f"{int(round(float(value))):,}",
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 5),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=chart_builder.body_size - 1,
            color=chart_builder.colors.get("leads", "#111111"),
        )

    for bar, value in zip(cost_bars, cost):
        if not value:
            continue
        ax_cost.annotate(
            _format_currency_label(float(value), currency_symbol, abbreviated=True),
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 5),
            textcoords="offset points",
            ha="center",
            va="bottom",
            fontsize=chart_builder.body_size - 1,
            color=chart_builder.colors.get("cost", "#D83A40"),
        )

    ax_leads.legend(
        [lead_bars[0], cost_bars[0]],
        ["Sales Leads", f"Cost ({currency_symbol})"],
        loc="upper left",
        fontsize=chart_builder.body_size,
        frameon=False,
    )
    plt.tight_layout()
    fig.savefig(out_path, dpi=180, facecolor="white")
    plt.close(fig)
    return out_path


def _build_combined_mix_chart(
    chart_builder: ChartBuilder, scope_key: str, mix_df: pd.DataFrame
) -> Path:
    out_path = chart_builder.charts_dir / f"{scope_key}.png"
    if (
        not isinstance(mix_df, pd.DataFrame)
        or mix_df.empty
        or "Campaign Type" not in mix_df.columns
        or float(mix_df.get("Cost", pd.Series(dtype=float)).fillna(0).sum()) <= 0
        and float(mix_df.get("Sales Leads", pd.Series(dtype=float)).fillna(0).sum()) <= 0
    ):
        return chart_builder._plot_empty_state(out_path, "No campaign mix data")

    chart_df = mix_df.copy()
    fig, axes = plt.subplots(2, 1, figsize=(3.2, 5.2))
    for ax, value_col, title in zip(
        axes,
        ("Cost", "Sales Leads"),
        ("Cost Share", "Lead Share"),
    ):
        source = chart_df[["Campaign Type", value_col]].copy()
        source[value_col] = pd.to_numeric(source[value_col], errors="coerce").fillna(0)
        source = source[source[value_col] > 0]
        if source.empty:
            ax.text(0.5, 0.5, "No data", ha="center", va="center")
            ax.axis("off")
            continue
        ax.pie(
            source[value_col],
            labels=source["Campaign Type"],
            autopct=lambda pct: f"{pct:.0f}%" if pct >= 4 else "",
            startangle=90,
            textprops={"fontsize": 7},
        )
        ax.set_title(title, fontsize=9, color="#666666")
        ax.axis("equal")
    plt.tight_layout()
    fig.savefig(out_path, dpi=220, facecolor="white")
    plt.close(fig)
    return out_path


def _load_source_generation_manifest(
    request_path: Path, artifact: Mapping[str, Any]
) -> dict[str, Any] | None:
    value = (artifact.get("source_files") or {}).get("source_generation_manifest")
    source_manifest_path = _resolve_path(value, request_path)
    if source_manifest_path and source_manifest_path.exists():
        return _read_json(source_manifest_path)
    fallback = request_path / "source_data" / "SOURCE_GENERATION_MANIFEST.json"
    if fallback.exists():
        return _read_json(fallback)
    return None


def _resolve_generated_file(
    source_manifest: Mapping[str, Any] | None, request_path: Path, key: str
) -> Path | None:
    if not source_manifest:
        return None
    value = (
        source_manifest.get("source_generation", {})
        .get("generated_files", {})
        .get(key)
    )
    path = _resolve_path(value, request_path)
    return path if path and path.exists() else None


def _resolve_manual_auction_path(request_path: Path) -> Path | None:
    auction_dir = request_path / "auction"
    if auction_dir.exists():
        for path in _csv_files(auction_dir):
            return path
    return None


def _resolve_manual_auction_sources(request_path: Path) -> dict[str, Path]:
    auction_dir = request_path / "auction"
    sources: dict[str, Path] = {}
    source_dirs = {
        "Google Ads": auction_dir / "google_ads",
        "Microsoft Ads": auction_dir / "microsoft_ads",
    }
    for label, source_dir in source_dirs.items():
        if source_dir.exists():
            for path in _csv_files(source_dir):
                sources[label] = path
                break
    return sources


def _csv_files(directory: Path) -> list[Path]:
    return sorted(
        path
        for path in directory.iterdir()
        if path.is_file() and path.suffix.lower() == ".csv"
    )


def _resolve_other_campaigns_dir(
    request_path: Path, source_manifest: Mapping[str, Any] | None
) -> Path | None:
    manual_dir = request_path / "other_campaigns"
    if manual_dir.exists() and any(manual_dir.glob("*.csv")):
        return manual_dir
    generated = _resolve_generated_file(source_manifest, request_path, "other_campaigns_dir")
    if generated and generated.exists():
        return generated
    return None


def _kpi_lookup(scope: Mapping[str, Any]) -> dict[str, Mapping[str, Any]]:
    return {
        str(item.get("key")): item
        for item in scope.get("kpis", [])
        if isinstance(item, Mapping) and item.get("key")
    }


def _kpi_value(kpis: Mapping[str, Mapping[str, Any]], metric: str) -> str:
    item = kpis.get(metric)
    return str(item.get("value") if item else "n/a")


def _kpi_yoy(kpis: Mapping[str, Mapping[str, Any]], metric: str) -> str:
    item = kpis.get(metric)
    return str(item.get("yoy_label") if item else "n/a")


def _localize_payload_currency(
    shape_text: dict[str, str],
    tables: dict[str, dict[str, Any]],
    currency_symbol: str,
) -> None:
    if currency_symbol == "£":
        return
    for object_id, value in list(shape_text.items()):
        shape_text[object_id] = _localize_currency_text(value, currency_symbol)
    for table_payload in tables.values():
        rows = table_payload.get("values")
        if not isinstance(rows, list):
            continue
        table_payload["values"] = [
            [_localize_currency_text(str(value), currency_symbol) for value in row]
            for row in rows
        ]


def _localize_currency_text(value: str, currency_symbol: str) -> str:
    return str(value).replace("£", currency_symbol) if currency_symbol != "£" else str(value)


def _maybe_set(
    shape_text: dict[str, str], ids: Sequence[str], index: int, value: str
) -> None:
    if index < len(ids):
        shape_text[str(ids[index])] = value


def _text_object_ids(presentation: dict[str, Any]) -> set[str]:
    object_ids: set[str] = set()
    for slide in presentation.get("slides") or []:
        for element in slide.get("pageElements") or []:
            if element.get("shape") and "text" in element.get("shape", {}):
                object_ids.add(str(element.get("objectId") or ""))
    return object_ids


def _format_quarter_subtitle(quarter: Any) -> str:
    return f"{quarter.label} ({quarter.start.strftime('%b')} - {quarter.end.strftime('%b %Y')})"


def _output_deck_title(client_name: str, period_label: str) -> str:
    generated_stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d %H%M UTC")
    return f"{client_name} {period_label} QBR - API Source Test - {generated_stamp}"


def _chart_scope_key(key: str) -> str:
    return f"native_qbr_{re.sub(r'[^a-z0-9]+', '_', key.lower()).strip('_')}"


def _fmt_currency(value: Any, currency_symbol: str = "£") -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{currency_symbol}{float(value):,.2f}"


def _format_currency_label(
    value: float, currency_symbol: str = "£", abbreviated: bool = False
) -> str:
    if abbreviated:
        absolute = abs(value)
        if absolute >= 1_000_000:
            return f"{currency_symbol}{value / 1_000_000:.1f}m"
        if absolute >= 1_000:
            return f"{currency_symbol}{value / 1_000:.1f}k"
        return f"{currency_symbol}{value:,.0f}"
    return f"{currency_symbol}{value:,.2f}"


def _format_currency_axis(value: float, currency_symbol: str = "£") -> str:
    absolute = abs(float(value))
    if absolute >= 1_000_000:
        return f"{currency_symbol}{float(value) / 1_000_000:.1f}m"
    if absolute >= 1_000:
        return f"{currency_symbol}{float(value) / 1_000:.0f}k"
    return f"{currency_symbol}{float(value):,.0f}"


def _fmt_percent(value: Any) -> str:
    if value is None or pd.isna(value):
        return "n/a"
    return f"{float(value) * 100:.2f}%"


__all__ = [
    "WWT_UK_QBR_TEMPLATE_MANIFEST",
    "build_wendy_wu_qbr_slides_payload",
    "generate_wendy_wu_qbr_google_slides",
]
