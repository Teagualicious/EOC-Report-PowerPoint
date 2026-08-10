<!-- Proposed by the project owner, 2026-07-15. Reviewed (critique, phasing,
     codebase integration map) in ../reviews/TEMPLATE_FIRST_MAPPER_REVIEW_2026-07-15.md
     — read BOTH before starting any implementation. Status: accepted direction,
     not yet scheduled. -->


# Template-First Mapper Architecture

## The Problem with the Current Approach

The current mapper operates by opening an existing client PowerPoint, traversing the shape tree, and performing in-place text/data replacements. This is fragile because:

- python-pptx must preserve every XML relationship, shape grouping, and formatting detail in the source file while modifications happen around them
- Slides authored in PowerPoint often contain deeply nested groups, phantom shapes, split text runs, and vendor-specific XML extensions that python-pptx doesn't fully model
- A single unexpected shape structure (grouped textbox, rotated chart, overlapping layers) can break the mapper at runtime with no graceful recovery
- Every new client template is a new source of edge cases

## The New Architecture: Ingest → Template → Map → Build

Instead of modifying existing files, the engine follows a four-phase pipeline:

```
[Client PPTX] → INGEST → [Template Schema (JSON)] → MAP → [Mapped Schema] → BUILD → [New PPTX]
```

The core insight is that **creating a new PPTX from a known schema is dramatically more reliable than modifying an unknown existing one**. The template schema acts as an intermediate representation (IR) that decouples the messy ingestion problem from the clean build problem.

---

## Phase 1: INGEST (PPTX → Template Schema)

### Purpose
Take a client's branded PowerPoint and extract a complete structural blueprint — every shape, every style property, every position — into a JSON template schema. This is a one-time operation per client template.

### What Gets Extracted Per Shape

For every shape on every slide, extract:

```json
{
  "shape_id": "slide3_shape7",
  "slide_index": 2,
  "shape_type": "text_box | chart | image | rectangle | line | group | table",
  "z_order": 7,
  "geometry": {
    "left": 914400,
    "top": 1524000,
    "width": 3657600,
    "height": 914400,
    "rotation": 0
  },
  "classification": "static | dynamic",
  "slot_name": null
}
```

Units are EMUs (English Metric Units) — python-pptx's native unit. 1 inch = 914400 EMU. Preserve these exactly; do not convert to inches/points during storage.

### Text Shape Extraction

For text boxes and text frames within other shapes, extract the full paragraph/run structure:

```json
{
  "shape_type": "text_box",
  "text_content": {
    "paragraphs": [
      {
        "alignment": "center",
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 1.0,
        "runs": [
          {
            "text": "X,XXX",
            "font": {
              "name": "Arial",
              "size_pt": 48,
              "bold": true,
              "italic": false,
              "color_hex": "#2E8B57",
              "underline": false
            }
          },
          {
            "text": "IMPRESSIONS",
            "font": {
              "name": "Arial",
              "size_pt": 18,
              "bold": true,
              "italic": false,
              "color_hex": "#1B2A4A",
              "underline": false
            }
          }
        ]
      }
    ]
  }
}
```

Key detail: text in PowerPoint is structured as paragraphs containing runs. Each run can have different formatting. A single text box with "X,XXX" in large green and "IMPRESSIONS" in small dark blue is two runs within one or two paragraphs. The ingestion must preserve this run-level granularity — it's what makes the rebuilt slide look identical to the original.

### Chart Extraction

For chart shapes, extract the chart type and visual configuration, not the data:

```json
{
  "shape_type": "chart",
  "chart_config": {
    "chart_type": "bar | column | line | pie | doughnut | area | scatter",
    "style": {
      "series_colors": ["#0054A6", "#2E8B57"],
      "has_legend": true,
      "legend_position": "bottom",
      "has_data_labels": true,
      "data_label_font": {"name": "Arial", "size_pt": 10, "color_hex": "#333333"},
      "category_axis": {
        "has_labels": true,
        "label_font": {"name": "Arial", "size_pt": 9}
      },
      "value_axis": {
        "has_labels": true,
        "min_value": null,
        "max_value": null
      }
    },
    "placeholder_categories": ["Cat 1", "Cat 2", "Cat 3", "Cat 4", "Cat 5"],
    "placeholder_values": [0, 0, 0, 0, 0]
  }
}
```

### Image/Logo Extraction

For images (logos, brand marks, decorative elements):

```json
{
  "shape_type": "image",
  "image_data": {
    "original_filename": "company_logo.png",
    "content_type": "image/png",
    "stored_path": "template_assets/{template_id}/company_logo.png",
    "crop": {
      "left": 0, "top": 0, "right": 0, "bottom": 0
    }
  }
}
```

The actual image bytes get saved to disk alongside the template schema. The schema references them by path.

### Decorative/Branding Shape Extraction

For rectangles, lines, dividers (like the blue bars in the example slide):

```json
{
  "shape_type": "rectangle",
  "fill": {
    "type": "solid | gradient | none",
    "color_hex": "#0054A6",
    "gradient_stops": null
  },
  "outline": {
    "color_hex": null,
    "width_pt": 0
  }
}
```

### Grouped Shapes

Groups are recursive. Extract as:

```json
{
  "shape_type": "group",
  "children": [
    { "shape_id": "...", "shape_type": "text_box", "..." : "..." },
    { "shape_id": "...", "shape_type": "rectangle", "..." : "..." }
  ]
}
```

Positions within groups are relative to the group's origin. Preserve that relationship.

### Slide-Level Properties

Per slide, also capture:

```json
{
  "slide_index": 0,
  "slide_layout_name": "Blank",
  "background": {
    "type": "solid | image | gradient",
    "color_hex": "#FFFFFF",
    "image_path": null
  },
  "shapes": [ "..." ]
}
```

### Presentation-Level Properties

At the top level of the template schema:

```json
{
  "template_id": "spectrum_reach_campaign_report_v1",
  "template_name": "Campaign Performance Report",
  "source_file": "original_template.pptx",
  "created_at": "2026-07-15T12:00:00Z",
  "slide_width_emu": 12192000,
  "slide_height_emu": 6858000,
  "slides": [ "..." ]
}
```

---

## Phase 2: CLASSIFY (Auto-Template)

### Purpose
After raw extraction, shapes need to be classified as **static** (branding elements reproduced exactly) or **dynamic** (data placeholders that get mapped).

### Auto-Classification Heuristics

The ingestion engine should apply these heuristics to suggest classifications:

| Signal | Classification | Reasoning |
|--------|---------------|-----------|
| Text contains "X,XXX", "XX%", "N/A", placeholder-like patterns | **dynamic** | Looks like a metric placeholder |
| Text contains client name, date patterns ("MONTH 1ST, 2026") | **dynamic** | Variable header fields |
| Shape is a chart | **dynamic** | Charts always receive mapped data |
| Shape is a table | **dynamic** | Tables always receive mapped data |
| Shape is a filled rectangle with no text | **static** | Decorative divider/branding bar |
| Shape is an image | **static** | Logo or brand image |
| Text is all-caps single word like "IMPRESSIONS", "CLICKS" | Depends on context | Could be a static label paired with a dynamic value — see paired detection below |

### Paired Element Detection

This is critical for KPI callouts like the one in the example image. The "X,XXX" (dynamic value) and "IMPRESSIONS" (static label) are often:

- Two runs in one text box (most common)
- Two separate text boxes positioned vertically adjacent
- Elements within a group

The classifier should detect these pairs and create a **compound slot**:

```json
{
  "slot_name": "kpi_impressions",
  "slot_type": "kpi_callout",
  "components": {
    "value": {
      "shape_ref": "slide3_shape7",
      "run_index": 0,
      "classification": "dynamic"
    },
    "label": {
      "shape_ref": "slide3_shape7",
      "run_index": 1,
      "classification": "static",
      "static_text": "IMPRESSIONS"
    }
  }
}
```

### User Review Step

After auto-classification, the GUI should present each slide visually with shapes highlighted in two colors (green for static, orange for dynamic). The user can:

- Override any classification
- Name the dynamic slots (e.g., rename "slide3_shape7_run0" → "total_impressions")
- Group related slots (e.g., "this chart and these 3 KPIs are all from the same data source")
- Mark shapes to exclude entirely

This reviewed classification gets saved as part of the final template schema.

---

## Phase 3: MAP (Data Source → Template Slots)

### Purpose
The user connects their data sources (Excel files, CSV exports, API responses) to the named dynamic slots in the template.

### Mapping Definition

```json
{
  "template_id": "spectrum_reach_campaign_report_v1",
  "data_source": {
    "type": "excel | csv | api",
    "path": "campaign_data.xlsx",
    "sheet": "Summary"
  },
  "mappings": [
    {
      "slot_name": "client_name",
      "source_column": "Client Name",
      "source_row": 0,
      "format": null
    },
    {
      "slot_name": "report_date",
      "source_column": null,
      "source_value": "auto:current_month",
      "format": "MMMM 1ST, YYYY"
    },
    {
      "slot_name": "total_impressions",
      "source_column": "Impressions",
      "source_row": "sum",
      "format": "#,###"
    },
    {
      "slot_name": "zip_code_chart",
      "chart_mapping": {
        "categories_column": "Zip Code",
        "values_column": "Delivered Impressions",
        "sort": "descending",
        "top_n": 5
      }
    }
  ]
}
```

### Mapping GUI

This is essentially what the current mapper GUI does, but instead of pointing at shape indices in a live PPTX, it points at named slots in a clean schema. The difference is:

- Slots have human-readable names ("total_impressions" not "slide 3 shape 7 run 0")
- The user sees a visual preview of the template with labeled slots
- Data source columns can be previewed alongside slot targets
- Validation can happen at map-time: "this slot expects a number but the column has text"

---

## Phase 4: BUILD (Mapped Schema → New PPTX)

### Purpose
Take the template schema + mapped data and construct a brand new PPTX from scratch. This is the phase where a different library strategy can be used.

### Build Strategy

**Option A: python-pptx Creation Mode (Recommended Starting Point)**

python-pptx is significantly more reliable when creating new presentations versus modifying existing ones. The build phase would:

1. Create a new `Presentation()` with the correct slide dimensions
2. For each slide in the template schema, add a blank slide
3. For each shape in the slide (respecting z-order), create the shape from scratch:
   - Text boxes: create with exact position/size, add paragraphs and runs with exact formatting
   - Charts: create with chart type, apply styling, inject mapped data
   - Images: add from stored asset files with exact position/crop
   - Rectangles/lines: create with exact fill/outline properties
4. Save as new file

**Option B: Template PPTX + Targeted Replacement (Hybrid)**

An alternative approach that captures more nuanced formatting:

1. During the INGEST phase, also save a "clean" copy of the original PPTX with all dynamic text replaced with well-known placeholder tokens (e.g., `{{total_impressions}}`)
2. At BUILD time, open this clean template and do simple string replacements of the tokens
3. For charts, use python-pptx to update chart data series

This hybrid preserves formatting that's hard to extract (shadows, 3D effects, complex gradients) but brings back some of the fragility of modifying existing files. However, since the tokens are well-known and the template is "clean" (no unexpected shapes), it's far more controlled.

**Option C: Direct OOXML Construction**

For maximum control, build the PPTX by writing the XML directly:

1. Maintain XML templates for each shape type (text box, chart, image, etc.) with Jinja2-style placeholders
2. Assemble the slide XML by combining shape templates
3. Package into the ZIP structure that is a PPTX file

This is the most work upfront but gives complete control and zero library dependency issues.

### Recommended Approach

Start with **Option A** (python-pptx creation mode). It handles 90% of cases well. For shapes where creation mode can't match the original fidelity (complex gradients, SmartArt, 3D effects), fall back to **Option B** for those specific shapes by keeping a reference copy of the original element's XML and splicing it in.

---

## Template Schema: Full Example

```json
{
  "schema_version": "1.0",
  "template_id": "spectrum_campaign_report_v1",
  "template_name": "Spectrum Reach Campaign Performance Report",
  "source_file": "Campaign_Report_Template.pptx",
  "created_at": "2026-07-15T12:00:00Z",
  "presentation": {
    "slide_width_emu": 12192000,
    "slide_height_emu": 6858000
  },
  "assets_directory": "templates/spectrum_campaign_report_v1/assets/",
  "slides": [
    {
      "slide_index": 0,
      "slide_label": "Campaign Overview",
      "background": {
        "type": "solid",
        "color_hex": "#FFFFFF"
      },
      "shapes": [
        {
          "shape_id": "s0_header_bar",
          "shape_type": "rectangle",
          "classification": "static",
          "z_order": 0,
          "geometry": {
            "left": 0,
            "top": 0,
            "width": 12192000,
            "height": 457200,
            "rotation": 0
          },
          "fill": {
            "type": "solid",
            "color_hex": "#0054A6"
          },
          "outline": {"width_pt": 0}
        },
        {
          "shape_id": "s0_client_date",
          "shape_type": "text_box",
          "classification": "dynamic",
          "slot_name": "header_client_date",
          "z_order": 1,
          "geometry": {
            "left": 228600,
            "top": 91440,
            "width": 8229600,
            "height": 365760,
            "rotation": 0
          },
          "text_content": {
            "paragraphs": [
              {
                "alignment": "left",
                "runs": [
                  {
                    "role": "dynamic",
                    "placeholder_text": "CLIENT NAME | MONTH 1ST, 2026",
                    "font": {
                      "name": "Arial",
                      "size_pt": 14,
                      "bold": true,
                      "color_hex": "#FFFFFF"
                    }
                  }
                ]
              }
            ]
          }
        },
        {
          "shape_id": "s0_section_title",
          "shape_type": "text_box",
          "classification": "static",
          "z_order": 2,
          "geometry": {
            "left": 228600,
            "top": 685800,
            "width": 8229600,
            "height": 548640,
            "rotation": 0
          },
          "text_content": {
            "paragraphs": [
              {
                "alignment": "left",
                "runs": [
                  {
                    "role": "static",
                    "text": "Extending Your Brand's Reach",
                    "font": {
                      "name": "Arial",
                      "size_pt": 28,
                      "bold": true,
                      "color_hex": "#1B2A4A"
                    }
                  }
                ]
              }
            ]
          }
        },
        {
          "shape_id": "s0_blue_divider_1",
          "shape_type": "rectangle",
          "classification": "static",
          "z_order": 3,
          "geometry": {
            "left": 0,
            "top": 1371600,
            "width": 12192000,
            "height": 91440,
            "rotation": 0
          },
          "fill": {
            "type": "solid",
            "color_hex": "#0054A6"
          }
        },
        {
          "shape_id": "s0_kpi_impressions",
          "shape_type": "text_box",
          "classification": "dynamic",
          "slot_name": "impressions_kpi",
          "slot_type": "kpi_callout",
          "z_order": 4,
          "geometry": {
            "left": 3657600,
            "top": 1600200,
            "width": 4876800,
            "height": 1143000,
            "rotation": 0
          },
          "text_content": {
            "paragraphs": [
              {
                "alignment": "center",
                "runs": [
                  {
                    "role": "dynamic",
                    "placeholder_text": "X,XXX",
                    "data_format": "#,###",
                    "font": {
                      "name": "Arial",
                      "size_pt": 48,
                      "bold": true,
                      "color_hex": "#2E8B57"
                    }
                  }
                ]
              },
              {
                "alignment": "center",
                "runs": [
                  {
                    "role": "static",
                    "text": "IMPRESSIONS",
                    "font": {
                      "name": "Arial",
                      "size_pt": 18,
                      "bold": true,
                      "color_hex": "#1B2A4A"
                    }
                  }
                ]
              }
            ]
          }
        },
        {
          "shape_id": "s0_zip_chart",
          "shape_type": "chart",
          "classification": "dynamic",
          "slot_name": "top_zip_codes_chart",
          "z_order": 6,
          "geometry": {
            "left": 228600,
            "top": 3200400,
            "width": 11735400,
            "height": 3200400,
            "rotation": 0
          },
          "chart_config": {
            "chart_type": "column",
            "title": "TOP 5 DELIVERED ZIP CODES",
            "title_font": {
              "name": "Arial",
              "size_pt": 14,
              "bold": true,
              "color_hex": "#1B2A4A"
            },
            "series": [
              {
                "name": "Impressions",
                "color_hex": "#0054A6",
                "data_label_font": {
                  "name": "Arial",
                  "size_pt": 10,
                  "color_hex": "#333333"
                }
              }
            ],
            "category_axis": {
              "label_font": {"name": "Arial", "size_pt": 9}
            },
            "value_axis": {
              "visible": false
            }
          }
        }
      ]
    }
  ],
  "slot_registry": {
    "header_client_date": {
      "type": "text",
      "description": "Client name and report date in header bar",
      "example": "ACME CORP | JULY 1ST, 2026"
    },
    "impressions_kpi": {
      "type": "number",
      "description": "Total impressions KPI callout",
      "format": "#,###"
    },
    "top_zip_codes_chart": {
      "type": "chart_data",
      "description": "Bar chart of top 5 zip codes by delivered impressions",
      "expects": {
        "categories": "list of zip code strings",
        "values": "list of impression counts"
      }
    }
  }
}
```

---

## Implementation Order

### Step 1: Schema Definition
Define the JSON schema formally (use JSON Schema or Pydantic models). This is the contract between all phases. Get this right first.

### Step 2: Ingestion Engine
Build the PPTX-to-schema extractor using python-pptx in read mode. This walks every slide, every shape, and populates the schema. Include the auto-classification heuristics.

### Step 3: Template Review GUI
Build the GUI screen where the user sees the extracted template, reviews classifications, names slots. This replaces the current mapper's "hunt through shapes" UX.

### Step 4: Mapping GUI
Adapt the current mapping interface to work with named slots instead of raw shape references. This should be a simpler, cleaner version of what exists.

### Step 5: Build Engine
Build the PPTX constructor that reads a mapped schema and produces a new file. Start with python-pptx creation mode.

### Step 6: Fidelity Testing
Compare original PPTX side-by-side with rebuilt PPTX. Identify gaps (gradients, shadows, effects that didn't survive the round-trip). Add fallback handling for those cases.

---

## Key Design Decisions

### Why JSON for the Template Schema
- Human-readable and debuggable
- Easy to diff between versions
- Claude Code and other tools can read/modify it
- Can be version-controlled in Git
- Pydantic models for validation

### Why EMUs for Positioning
- Native unit of OOXML — no conversion rounding errors
- python-pptx uses EMUs internally
- Exact reproduction of original positions guaranteed

### Why Separate Assets Directory
- Images stored as actual files, not base64 in JSON
- Template schema stays small and readable
- Assets can be swapped (e.g., update a logo) without re-ingesting

### Why a Slot Registry
- Single source of truth for what data the template expects
- Enables validation: "you mapped a text column to a number slot"
- Enables documentation: each slot has a human description
- Enables reuse: different data sources can map to the same template

---

## Edge Cases to Handle

1. **SmartArt**: python-pptx cannot read SmartArt. Detect and warn the user; offer to screenshot and embed as a static image.
2. **Embedded videos/audio**: Detect and flag as unsupported; copy as static placeholder.
3. **Master slide / layout inheritance**: Some formatting comes from the slide master, not the shape itself. The ingestion must resolve inherited fonts/colors.
4. **Theme colors**: PowerPoint often stores colors as theme references ("Accent 1") not hex values. Resolve these to actual hex during ingestion by reading the theme XML.
5. **Tables**: Extract row/column structure, cell merges, per-cell formatting. Tables are complex but important.
6. **Connectors and lines**: Preserve start/end points and routing.
7. **Text autofit**: PowerPoint can auto-shrink text to fit a box. Record the autofit setting so the builder can replicate it.

---

## File Structure

```
ingestion_engine/
├── templates/
│   ├── schema.py              # Pydantic models for template schema
│   ├── ingester.py            # Phase 1: PPTX → raw schema
│   ├── classifier.py          # Phase 2: auto-classify static/dynamic
│   ├── mapper.py              # Phase 3: data source → slot mapping
│   └── builder.py             # Phase 4: mapped schema → new PPTX
├── template_store/
│   ├── {template_id}/
│   │   ├── template.json      # The template schema
│   │   ├── mapping.json       # Saved mapping configuration
│   │   └── assets/            # Extracted images, logos
├── gui/
│   ├── template_review.py     # Template review/classification GUI
│   └── slot_mapper.py         # Data-to-slot mapping GUI
```
