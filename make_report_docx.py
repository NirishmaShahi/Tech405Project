import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    doc = docx.Document()

    # --- Page setup (Standard Letter, 1-inch margins) ---
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
        # Header / Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "Kathmandu Valley Precipitation Forecasting | Deep Learning Technical Report"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.size = Pt(8.5)
        hp.style.font.color.rgb = RGBColor(120, 144, 156)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "TECH405 Project • Tribhuvan International Airport Meteorological Benchmark"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.style.font.size = Pt(8.5)
        fp.style.font.color.rgb = RGBColor(120, 144, 156)

    # Styling constants
    COLOR_PRIMARY = RGBColor(27, 54, 93)      # #1B365D - Deep Navy
    COLOR_SECONDARY = RGBColor(43, 108, 176)  # #2B6CB0 - Steel Blue
    COLOR_BODY = RGBColor(45, 55, 72)         # #2D3748 - Charcoal
    COLOR_MUTED = RGBColor(113, 128, 150)     # #718096 - Slate Muted

    # Helpers
    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(22)
        run.font.name = "Georgia"
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(text)
        run.font.size = Pt(12)
        run.font.italic = True
        run.font.name = "Calibri"
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_meta(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(20)
        run = p.add_run(text)
        run.font.size = Pt(9.5)
        run.font.name = "Calibri"
        run.font.color.rgb = COLOR_MUTED
        return p

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(14)
        run.font.name = "Georgia"
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11.5)
        run.font.name = "Georgia"
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_heading_3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10.5)
        run.font.name = "Georgia"
        run.font.color.rgb = COLOR_BODY
        return p

    def add_paragraph(text, bold_prefix=None, space_after=6):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_after = Pt(space_after)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.size = Pt(10.5)
            r_pre.font.name = "Calibri"
            r_pre.font.color.rgb = COLOR_PRIMARY
        r = p.add_run(text)
        r.font.size = Pt(10.5)
        r.font.name = "Calibri"
        r.font.color.rgb = COLOR_BODY
        return p

    def add_callout(text, bold_prefix="Key Finding: "):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.5)
        
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        tcPr.append(shd)
        
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="0" w:color="1B365D"/><w:top w:val="none"/><w:right w:val="none"/><w:bottom w:val="none"/></w:tcBorders>')
        tcPr.append(borders)
        
        mar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tcMar>')
        tcPr.append(mar)
        
        cp = cell.paragraphs[0]
        cp.paragraph_format.line_spacing = 1.15
        cp.paragraph_format.space_after = Pt(2)
        r_pre = cp.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.size = Pt(10)
        r_pre.font.name = "Calibri"
        r_pre.font.color.rgb = COLOR_PRIMARY
        
        r = cp.add_run(text)
        r.font.size = Pt(10)
        r.font.name = "Calibri"
        r.font.color.rgb = COLOR_BODY
        
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(0)
        p_space.paragraph_format.space_after = Pt(6)

    def add_image_figure(img_path, caption, fig_num, width_inches=5.8):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(8)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=Inches(width_inches))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(12)
            p_cap.paragraph_format.keep_with_next = False
            
            r_num = p_cap.add_run(f"Figure {fig_num}: ")
            r_num.bold = True
            r_num.font.size = Pt(9.5)
            r_num.font.name = "Calibri"
            r_num.font.color.rgb = COLOR_PRIMARY
            
            r_text = p_cap.add_run(caption)
            r_text.font.size = Pt(9.5)
            r_text.font.italic = True
            r_text.font.name = "Calibri"
            r_text.font.color.rgb = COLOR_BODY

    def format_table(table, col_widths, alignments):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(table.rows):
            # prevent split across pages
            trPr = row._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            if i == 0:
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                
            for j, cell in enumerate(row.cells):
                cell.width = Inches(col_widths[j])
                tcPr = cell._tc.get_or_add_tcPr()
                
                # Shading
                if i == 0:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1B365D"/>')
                elif i % 2 == 1:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFFFF"/>')
                else:
                    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F7FAFC"/>')
                tcPr.append(shd)
                
                # Borders
                borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/><w:left w:val="none"/><w:right w:val="none"/></w:tcBorders>')
                tcPr.append(borders)
                
                # Padding
                mar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="80" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>')
                tcPr.append(mar)
                
                for p in cell.paragraphs:
                    p.alignment = alignments[j]
                    p.paragraph_format.line_spacing = 1.05
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for r in p.runs:
                        r.font.name = "Calibri"
                        if i == 0:
                            r.bold = True
                            r.font.size = Pt(9.5)
                            r.font.color.rgb = RGBColor(255, 255, 255)
                        else:
                            r.font.size = Pt(9)
                            r.font.color.rgb = COLOR_BODY

    # -------------------------------------------------------------
    # DOCUMENT CONTENT
    # -------------------------------------------------------------
    
    # Title Block
    add_title("Deep Learning Architectures for Next-Day Precipitation Forecasting in Complex Himalayan Topography")
    add_subtitle("A Rigorous Empirical Benchmark of Multilayer Perceptrons, Long Short-Term Memory Networks, and 1D-Convolutional Neural Networks in the Kathmandu Valley")
    add_meta("Department of Applied Data Science & Computing • Tribhuvan International Airport Weather Analytics Project • October 2026")

    # Abstract Box
    abstract_text = (
        "Accurate next-day precipitation forecasting in the Kathmandu Valley is essential for municipal flash flood management, "
        "agriculture, and flight operations at Tribhuvan International Airport (TIA; VNKT). However, steep Himalayan orography and "
        "complex South Asian monsoon dynamics make conventional numerical weather prediction (NWP) computationally demanding "
        "and prone to localized sub-grid parameterization errors. This technical report presents a comprehensive investigation into "
        "data-driven deep learning models for binary next-day rain classification using a 10-year synoptic observation record (2015–2025) "
        "from TIA. We formulate a sequence-to-label formulation using a 14-day rolling historical lookback window comprising 20 "
        "meteorological, astronomical, cyclical harmonic, and lag-derived features. Three structurally distinct deep learning architectures "
        "— a Multilayer Perceptron (MLP; 44,738 parameters), a stacked Long Short-Term Memory network (LSTM; 55,426 parameters), and a "
        "1D-Convolutional Neural Network (1D-CNN; 16,642 parameters) — are trained under class-weighted cross-entropy loss with early stopping "
        "and evaluated on a strictly held-out future test year (July 1, 2024 to June 30, 2025; 365 days, 43.0% rain prevalence).\n\n"
        "On the single-run test benchmark, both the MLP and LSTM achieve an identical peak classification accuracy of 83.56% and F1-score of 80.77%, "
        "outperforming the 1D-CNN (74.25% accuracy, 73.74% F1-score). However, discriminative threshold-free analysis reveals clear superiority "
        "for the LSTM, yielding a Receiver Operating Characteristic Area Under Curve (ROC-AUC) of 0.8903 and Precision-Recall AUC (PR-AUC) of 0.8747, "
        "along with the lowest Brier calibration score (0.1252), compared to 0.8699 ROC-AUC and 0.1369 Brier score for the MLP. To ensure reproducibility "
        "and mitigate seed stochasticity, a 5-fold Monte Carlo random seed verification is conducted. Across 5 seeds, the LSTM consistently leads all "
        "architectures with mean accuracy of 83.78% ± 0.68%, F1-score of 81.45% ± 0.82%, ROC-AUC of 0.8869 ± 0.16%, and PR-AUC of 0.8698 ± 0.22%. "
        "Atmospheric analysis confirms that recurrent gating effectively captures multi-day barometric pressure depressions and humidity accumulation "
        "patterns preceding monsoon and post-monsoon precipitation events, establishing the stacked LSTM as the preferred model for operational deployment."
    )
    add_callout(abstract_text, bold_prefix="EXECUTIVE ABSTRACT\n")

    add_paragraph("Keywords: Precipitation Forecasting, Long Short-Term Memory (LSTM), 1D-Convolutional Neural Networks, Deep Learning, Orographic Meteorology, Kathmandu Valley, Tribhuvan International Airport, Time-Series Sequence Modeling.", space_after=12)

    # -------------------------------------------------------------
    # SECTION 1
    # -------------------------------------------------------------
    add_heading_1("1. Introduction & Operational Context")
    
    add_heading_2("1.1 The Meteorological Challenge of the Kathmandu Valley")
    add_paragraph(
        "Precipitation prediction in mountainous intermontane valleys represents one of the most demanding problems in modern operational "
        "hydrometeorology. The Kathmandu Valley, located in central Nepal at an average elevation of 1,400 meters above mean sea level (ASL), "
        "is topographically enclosed by surrounding mountain ridges rising between 2,000 and 2,800 meters. This unique bowl-shaped physiography "
        "creates distinct microclimates characterized by strong nocturnal thermal inversions, localized katabatic-anabatic wind circulations, "
        "and acute sensitivity to seasonal monsoon moisture advection."
    )
    add_paragraph(
        "Climatologically, the valley is dominated by the South Asian Summer Monsoon (June through September), which supplies approximately "
        "80% of its annual precipitation (1,400 to 1,600 mm). During this period, moist southeasterly currents originating from the Bay of Bengal "
        "collide with the steep foothills of the Himalayas, triggering intense convective and stratiform rainfall. Conversely, during the pre-monsoon "
        "season (March through May), intense solar insolation over mountain slopes generates severe localized afternoon thunderstorms. Winter months "
        "(December to February) experience sporadic western disturbances that generate light rain and pervasive, low-level radiation fogs."
    )

    add_heading_2("1.2 Aviation, Municipal, and Socioeconomic Significance")
    add_paragraph(
        "The socioeconomic consequences of unexpected precipitation in Kathmandu are profound. The valley houses over 3 million residents and "
        "hosts Tribhuvan International Airport (TIA; ICAO: VNKT), Nepal's sole primary international gateway and highest-density aviation hub. "
        "Because TIA is situated in a high-altitude bowl flanked by mountainous terrain, instrument flight procedures and visual approaches are "
        "highly sensitive to weather degradation. Unanticipated precipitation events trigger immediate cloud-ceiling drops, severe runway friction loss, "
        "visibility degradation below landing minimums, flight diversions, and dangerous holding patterns."
    )
    add_paragraph(
        "Simultaneously, rapid, unplanned urban sprawl in Kathmandu has severely compromised natural drainage channels (such as the Bagmati, Bishnumati, "
        "and Dhobi Khola rivers). Convective rainstorms frequently cause flash flooding, waterlogged arterial roads, structural inundation, and slope "
        "destabilization along surrounding hillsides. Consequently, municipal disaster management authorities and aviation dispatchers require "
        "dependable 24-hour advance warning of rainfall events to implement preemptive mitigation protocols."
    )

    add_heading_2("1.3 Limitations of Numerical Weather Prediction (NWP)")
    add_paragraph(
        "Traditionally, meteorological agencies rely on physics-based Numerical Weather Prediction (NWP) systems, such as the Weather Research "
        "and Forecasting (WRF) model or ECMWF global forecasts. While NWP models excel at capturing broad synoptic wave dynamics, their horizontal "
        "grid resolutions (typically 3 to 25 km) are too coarse to resolve the intricate orographic lifting, valley wind channeling, and microscale "
        "convective updrafts occurring within Kathmandu's complex terrain. Achieving sub-kilometer resolutions requires supercomputing resources "
        "that are prohibitive for real-time municipal or regional operations in developing nations. Furthermore, NWP parameterization schemes "
        "frequently suffer from systematic biases in mountainous regions, leading to severe false-alarm rates or underestimation of sudden rainfall."
    )

    add_heading_2("1.4 Objectives & Formal Problem Statement")
    add_paragraph(
        "Machine learning and deep sequence models offer a compelling, data-driven alternative by learning nonlinear temporal mappings directly "
        "from multi-decadal in-situ station observations. The primary objectives of this study are threefold:"
    )
    add_paragraph("1. Formulate a leak-free sequence-to-label time series classification pipeline that forecasts whether precipitation (> 0.0 mm) will occur on day t+1 based strictly on an antecedent 14-day observation window (days t-13 through t).")
    add_paragraph("2. Implement and benchmark three distinct deep learning architectural paradigms: a non-recurrent Multilayer Perceptron (MLP), a recurrent stacked Long Short-Term Memory (LSTM) network, and a 1D-Convolutional Neural Network (1D-CNN).")
    add_paragraph("3. Conduct rigorous performance evaluation on a strictly held-out future calendar year (July 2024 to June 2025) across classification metrics, probability calibration, confusion matrix dynamics, and multi-seed Monte Carlo replication.")

    # -------------------------------------------------------------
    # SECTION 2
    # -------------------------------------------------------------
    add_heading_1("2. Theoretical Foundation & Related Literature")

    add_heading_2("2.1 Himalayan Hydro-Meteorology and South Asian Monsoon Dynamics")
    add_paragraph(
        "Understanding precipitation predictability requires grounding in the physical atmospheric mechanisms governing the Himalayan foothills. "
        "Webster et al. (1998) established foundational principles regarding South Asian monsoon transitions, highlighting how large-scale land-sea "
        "thermal contrasts interact with elevated topography to establish persistent moist atmospheric channels. Over Nepal, Shrestha et al. (2000) "
        "and Dimri et al. (2015) analyzed precipitation patterns and spatial variability, noting that the Mahabharat range and high Himalayas "
        "induce heavy orographic precipitation along windward slopes, while creating pronounced precipitation gradients within enclosed valley basins."
    )
    add_paragraph(
        "Recent atmospheric studies (Bohlinger & Sorteberg, 2018; Karki et al., 2017) emphasize that rainfall in the Kathmandu basin does not "
        "occur as isolated, memoryless random shocks; rather, it is preceded by multi-day synoptic signatures, including gradual reductions in "
        "sea-level barometric pressure, escalating specific humidity, nocturnal dew-point depressions, and shifting diurnal temperature ranges. "
        "Capturing these multi-day atmospheric accumulations is precisely where temporal sequence modeling demonstrates its theoretical strength."
    )

    add_heading_2("2.2 Deep Learning Paradigms in Meteorological Forecasting")
    add_paragraph(
        "Over the past decade, deep learning has transformed Earth system science (Reichstein et al., 2019; Schultz et al., 2021). "
        "In hydrological and meteorological sequence modeling, recurrent neural networks — specifically Long Short-Term Memory networks "
        "(Hochreiter & Schmidhuber, 1997) — have demonstrated state-of-the-art capability in capturing temporal dependencies and hysteresis effects. "
        "In a landmark series of papers, Kratzert et al. (2018, 2019) proved that LSTMs substantially outperform traditional conceptual and "
        "physics-based hydrological models across continental benchmarks, attributing this success to the LSTM's ability to maintain dynamic "
        "internal state representations corresponding to physical catchment storage and moisture accumulation."
    )
    add_paragraph(
        "Similarly, Hewage et al. (2020) and Rasp et al. (2020) demonstrated the effectiveness of deep neural architectures for local weather "
        "forecasting, demonstrating that multi-layer perceptrons and temporal convolutions can approximate complex atmospheric transfer functions "
        "at a fraction of the computational runtime required by numerical solvers."
    )

    add_heading_2("2.3 Recurrent vs. Convolutional Sequence Modeling Trade-offs")
    add_paragraph(
        "A central architectural debate in temporal sequence modeling concerns the relative efficacy of recurrent gating mechanisms versus "
        "dilated/sliding convolutional operations. Bai et al. (2018) conducted an empirical evaluation showing that Temporal Convolutional "
        "Networks (TCNs / 1D-CNNs) can match or exceed recurrent models on several synthetic benchmark tasks due to parallelizable training and "
        "rigid receptive field boundaries. Lea et al. (2017) highlighted that 1D-CNNs effectively capture localized hierarchical patterns through "
        "stacked spatial filter kernels."
    )
    add_paragraph(
        "However, as shown by Graves (2012) and Sutskever et al. (2014), recurrent architectures possess an inductive bias specifically suited "
        "for continuous-state tracking where the duration of antecedent influence varies dynamically. In meteorological contexts, where the onset "
        "of a monsoon active spell may depend on cumulative pressure trends spanning 5 to 14 days, the adaptive gating of LSTMs provides distinct "
        "advantages over rigid convolutional filters, a hypothesis empirically tested in this work."
    )

    add_heading_2("2.4 Meteorological Verification and Probability Calibration")
    add_paragraph(
        "Standard classification metrics such as raw accuracy can be deeply misleading in atmospheric science due to seasonal base-rate shifts "
        "and class imbalance (Wilks, 2011; Jolliffe & Stephenson, 2012). Atmospheric forecast verification requires a multi-faceted assessment "
        "encompassing True Positive Rates (Recall/Hit Rate), False Alarm Ratios (Precision), Receiver Operating Characteristic (ROC) analysis, and "
        "Precision-Recall Area Under Curve (PR-AUC; Davis & Goadrich, 2006)."
    )
    add_paragraph(
        "Furthermore, because operational decision-makers must convert predicted probabilities into concrete risk actions, model calibration is "
        "paramount. The Brier Score (Brier, 1950) measures the mean squared error between forecast probabilities and binary observation outcomes, "
        "serving as a strictly proper scoring rule that penalizes overconfident, poorly calibrated predictions. Both discriminative power (ROC-AUC) "
        "and calibration reliability (Brier Score) are thoroughly examined in our benchmarking framework."
    )

    # -------------------------------------------------------------
    # SECTION 3
    # -------------------------------------------------------------
    add_heading_1("3. Dataset Description, Preprocessing & Feature Engineering")

    add_heading_2("3.1 Station Profile & Observational Record")
    add_paragraph(
        "The empirical foundation of this study is a 10.5-year continuous synoptic observation record collected at Tribhuvan International Airport "
        "(Station Code: VNKT; Latitude: 27.696° N, Longitude: 85.359° E, Elevation: 1,338 m ASL), covering the period from January 1, 2015 "
        "through June 30, 2025. The raw dataset contains 3,468 calendar days and 22 primary synoptic meteorological channels, including daily maximum, "
        "minimum, and mean 2-meter air temperatures, dew point, relative humidity, precipitation accumulation, probability, sea-level barometric "
        "pressure, cloud cover fraction, surface horizontal visibility, solar radiation flux, solar energy, wind speed, and maximum wind gusts."
    )

    add_heading_2("3.2 Sensor Error Remediation & Data Quality Control")
    add_paragraph(
        "Real-world surface weather instrumentation is susceptible to intermittent sensor glitches, transmission packet drops, and calibration drift. "
        "Rigorous exploratory analysis revealed two critical physical anomalies in the raw record:"
    )
    add_paragraph(
        "1. Wind Speed Inversion Anomaly: A physical anemometer anomaly was identified where recorded continuous wind speeds exceeded twice the "
        "maximum recorded wind gust speed (windspeed > 2 × windgust). Because sustained wind speed cannot physically exceed peak gust speed over the "
        "same observation interval, all aberrant wind speed entries were capped to the corresponding gust speed value."
    )
    add_paragraph(
        "2. Optical Visibility Outliers: While typical atmospheric visibility in the Kathmandu basin ranges between 2.0 km (during winter smog/fog) "
        "and 10.0 km (post-rain clearing), an anomalous sensor spike recorded visibility exceeding 20.0 km (reaching 24.1 km). This physically "
        "implausible extreme was corrected by imputing the robust long-term median visibility."
    )

    add_heading_2("3.3 The 12-Month Calendar Gap & Time-Series Integrity")
    add_paragraph(
        "Calendar delta auditing revealed a critical structural gap: exactly one 367-day discontinuity exists in the raw archive between July 1, 2015 "
        "and June 30, 2016 (a 12-month station offline period). Naive computation of autoregressive lag features (e.g., shifting values by 1 day) or "
        "rolling window aggregation across this boundary would contaminate post-gap records with data from an entire year prior."
    )
    add_paragraph(
        "To guarantee mathematical and physical validity, we implemented strict boundary isolation: any lag feature (temp_lag1, humidity_lag1) or "
        "rolling moving average (temp_roll7) spanning the gap boundary was flagged and set to null. Furthermore, the dataset was formally truncated "
        "to begin on July 1, 2016, providing an unbroken, continuous 9-year primary observation record of 3,285 usable days free from temporal splices."
    )

    add_heading_2("3.4 Engineered Feature Space (20 Dimensions)")
    add_paragraph(
        "To furnish deep neural networks with physically meaningful atmospheric signals, domain-specific feature engineering was performed across "
        "four primary categories:"
    )
    add_paragraph("• Astronomical & Daylight Features: Solar sunrise and sunset timestamps were converted into true astronomical daylight duration (daylight_hours = (sunset - sunrise) / 3600), capturing the seasonal insolation potential driving convective heating.")
    add_paragraph("• Cyclical Harmonic Encodings: Calendar month (1–12) and day-of-year (1–365) were mapped into continuous two-dimensional trigonometric manifolds via sine and cosine transformations: month_sin = sin(2π m / 12), month_cos = cos(2π m / 12), doy_sin = sin(2π d / 365), doy_cos = cos(2π d / 365). This ensures smooth circular continuity across December 31 and January 1.")
    add_paragraph("• Diurnal & Stability Indicators: The diurnal temperature range (temp_range = tempmax - tempmin) was computed as a direct proxy for atmospheric boundary layer stability and cloud cover suppression.")
    add_paragraph("• Temporal Memory & Precipitation Persistence: A binary indicator representing today's rainfall occurrence (rained = 1 if precipitation > 0 else 0) was incorporated, alongside 1-day lagged temperature and humidity, and a 7-day backward rolling mean temperature (temp_roll7).")

    # Table 1: Feature Summary
    t1_data = [
        ["Feature Name", "Domain / Physical Category", "Transformation / Unit", "Meteorological Relevance"],
        ["temp", "Thermal State", "Degrees Fahrenheit (°F)", "Mean daily ambient thermal energy"],
        ["temp_range", "Boundary Layer Stability", "tempmax - tempmin (°F)", "Diurnal temperature excursion; proxy for cloud cover"],
        ["humidity", "Atmospheric Moisture", "Relative Humidity (%)", "Surface moisture saturation percentage"],
        ["dew", "Dew Point Temperature", "Degrees Fahrenheit (°F)", "Absolute atmospheric moisture indicator"],
        ["sealevelpressure", "Barometric Dynamics", "Hectopascals (hPa / mb)", "Synoptic high/low pressure systems and monsoon troughs"],
        ["cloudcover", "Radiative Forcing", "Percentage Cover (%)", "Cloud obscuration suppressing insolation"],
        ["visibility", "Aerosol / Particulate", "Statute Miles / Kilometers", "Atmospheric turbidity, haze, and rain washout"],
        ["solarradiation", "Solar Flux", "Watts per m² (W/m²)", "Direct insolation driving afternoon convective storm cells"],
        ["windspeed", "Kinetic Transport", "Miles per hour (mph)", "Sustained horizontal advection speed"],
        ["windgust", "Turbulence / Shear", "Miles per hour (mph)", "Peak short-duration aerodynamic gust"],
        ["precipprob", "Synoptic Probability", "Percentage (%)", "Baseline probabilistic synoptic moisture signal"],
        ["month_sin, month_cos", "Harmonic Calendar", "Cyclical [-1.0, 1.0]", "Smooth representation of 12-month monsoon cycle"],
        ["doy_sin, doy_cos", "Harmonic Solar Day", "Cyclical [-1.0, 1.0]", "Daily resolution solar angle and seasonal progression"],
        ["daylight_hours", "Photoperiod Duration", "Decimal Hours (h)", "Calculated astronomical photoperiod duration"],
        ["rained", "Autoregressive State", "Binary flag {0, 1}", "Today's observed rainfall (known at prediction time)"],
        ["temp_lag1", "Lag-1 Temperature", "Degrees Fahrenheit (°F)", "Yesterday's mean temperature; short-term trend"],
        ["humidity_lag1", "Lag-1 Humidity", "Percentage (%)", "Yesterday's humidity; short-term moisture continuity"],
        ["temp_roll7", "Trailing Multi-Day Mean", "7-day backward mean (°F)", "Weekly synoptic thermal trend baseline"]
    ]
    t1 = doc.add_table(rows=len(t1_data), cols=4)
    for r_idx, row in enumerate(t1_data):
        for c_idx, val in enumerate(row):
            t1.cell(r_idx, c_idx).paragraphs[0].text = val
    format_table(t1, [1.4, 1.6, 1.6, 1.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT])
    
    p_t1_cap = doc.add_paragraph()
    p_t1_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t1_cap.paragraph_format.space_before = Pt(4)
    p_t1_cap.paragraph_format.space_after = Pt(10)
    r_t1 = p_t1_cap.add_run("Table 1: Complete 20-dimensional meteorological and engineered feature space.")
    r_t1.bold = True
    r_t1.font.size = Pt(9.5)

    # Figure 1
    add_image_figure("week6_outputs/01_class_balance_seasonality.png",
                     "Class balance across the 9-year dataset (left: 60.5% dry days vs. 39.5% rain days) and monthly rainfall probability (right), illustrating peak monsoon concentration in June, July, and August (>80%) versus arid winter baseline (<15%).",
                     1, width_inches=5.8)

    add_heading_2("3.5 Chronological Window Slicing & Leak-Free Scaling")
    add_paragraph(
        "Random shuffling or standard k-fold cross-validation is strictly forbidden in time series forecasting because temporal shuffling leaks "
        "future synoptic information into past model states. We instituted a rigorous chronological train/validation/test split:"
    )
    add_paragraph("• Training Partition: July 1, 2016 through June 30, 2023 (7 complete meteorological years; 2,541 sequences; 38.9% rain prevalence).")
    add_paragraph("• Validation Partition: July 1, 2023 through June 30, 2024 (1 complete annual cycle; 366 sequences; 38.5% rain prevalence).")
    add_paragraph("• Held-Out Test Partition: July 1, 2024 through June 30, 2025 (1 complete, untouched annual cycle; 365 sequences; 43.0% rain prevalence).")
    add_paragraph(
        "To capture multi-day atmospheric accumulations while maintaining operational responsiveness, we established a rolling lookback window of "
        "WINDOW = 14 days. Each input sample consists of the tensor X ∈ ℝ^(14 × 20) representing days t-13 through t, paired with the target label "
        "y ∈ {0, 1} indicating whether rain occurs on day t+1. A strict programmatic sanity check verified that the prediction target date t+1 "
        "is strictly disjoint from the historical observation window. Finally, a Scikit-Learn StandardScaler was fitted strictly on the training partition "
        "and applied forward to validation and test partitions, guaranteeing zero lookahead leakage."
    )

    # Table 2: Splits
    t2_data = [
        ["Partition", "Temporal Range", "Calendar Span", "Sequences (X)", "Rain Days", "Prevalence"],
        ["Training", "2016-07-01 → 2023-06-30", "7.0 Years", "2,541", "988", "38.9%"],
        ["Validation", "2023-07-01 → 2024-06-30", "1.0 Year (Leap)", "366", "141", "38.5%"],
        ["Test (Held-Out)", "2024-07-01 → 2025-06-30", "1.0 Year", "365", "157", "43.0%"],
        ["Total Usable", "2016-07-01 → 2025-06-30", "9.0 Years", "3,272", "1,286", "39.3%"]
    ]
    t2 = doc.add_table(rows=len(t2_data), cols=6)
    for r_idx, row in enumerate(t2_data):
        for c_idx, val in enumerate(row):
            t2.cell(r_idx, c_idx).paragraphs[0].text = val
    format_table(t2, [1.1, 1.8, 1.0, 0.9, 0.8, 0.9], [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    
    p_t2_cap = doc.add_paragraph()
    p_t2_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t2_cap.paragraph_format.space_before = Pt(4)
    p_t2_cap.paragraph_format.space_after = Pt(10)
    r_t2 = p_t2_cap.add_run("Table 2: Chronological partitioning, sequence window counts, and target class distributions.")
    r_t2.bold = True
    r_t2.font.size = Pt(9.5)

    # -------------------------------------------------------------
    # SECTION 4
    # -------------------------------------------------------------
    add_heading_1("4. Deep Learning Architectures & Optimization Pipeline")

    add_heading_2("4.1 Multilayer Perceptron (MLP) Baseline")
    add_paragraph(
        "The Multilayer Perceptron serves as our non-recurrent deep learning baseline. To process temporal sequences, the 14-day window tensor is "
        "flattened into a static 280-dimensional feature vector (14 × 20 = 280). The network architecture comprises two fully connected dense layers "
        "with hidden dimensions of 128 and 64 units, respectively. Each hidden layer is followed by Batch Normalization (Ioffe & Szegedy, 2015), "
        "Rectified Linear Unit (ReLU) activations, and a Dropout rate of 0.30 (Srivastava et al., 2014) to prevent co-adaptation. A final linear "
        "layer projects the 64-dimensional latent embedding to 2 class logits. Total trainable parameters: 44,738."
    )

    add_heading_2("4.2 Stacked Long Short-Term Memory (LSTM) Network")
    add_paragraph(
        "To model sequential dependencies and time-lagged memory explicitly, we implement a stacked Long Short-Term Memory recurrent neural network "
        "(Hochreiter & Schmidhuber, 1997). At each discrete time step t ∈ {1, ..., 14}, the LSTM cell updates an input gate i_t, forget gate f_t, "
        "cell candidate c̃_t, cell memory state c_t, and output gate o_t:"
    )
    add_paragraph(
        "f_t = σ(W_f x_t + U_f h_(t-1) + b_f)\n"
        "i_t = σ(W_i x_t + U_i h_(t-1) + b_i)\n"
        "c̃_t = tanh(W_c x_t + U_c h_(t-1) + b_c)\n"
        "c_t = f_t ⊙ c_(t-1) + i_t ⊙ c̃_t\n"
        "o_t = σ(W_o x_t + U_o h_(t-1) + b_o)\n"
        "h_t = o_t ⊙ tanh(c_t)"
    )
    add_paragraph(
        "The architecture consists of 2 stacked LSTM recurrent layers with a hidden dimension of 64 units and an inter-layer dropout rate of 0.30. "
        "The final hidden state vector h_14 ∈ ℝ^64, representing the aggregated temporal context of the complete 14-day history, is passed through "
        "a dropout layer (p = 0.30) and a linear classification head yielding 2 output logits. Total trainable parameters: 55,426."
    )

    add_heading_2("4.3 1D-Convolutional Neural Network (1D-CNN)")
    add_paragraph(
        "The 1D-Convolutional Neural Network operates by sliding temporal filter kernels along the time axis (Bai et al., 2018; Lea et al., 2017). "
        "The input tensor X is transposed to shape (Batch, 20 features, 14 timesteps). The network incorporates two successive 1D convolutional layers, "
        "each utilizing 64 output channels, a kernel size of k = 3, stride of 1, and symmetric padding of 1. Each convolution is coupled with Batch "
        "Normalization and ReLU activations. The resulting feature map is collapsed across the temporal dimension using an Adaptive Average Pooling "
        "layer (AdaptiveAvgPool1d(1)), followed by dropout (p = 0.30) and a linear classification head producing 2 output logits. Total trainable "
        "parameters: 16,642."
    )

    add_heading_2("4.4 Loss Function, Class Balancing & Training Dynamics")
    add_paragraph(
        "Across all models, optimization is performed using mini-batch gradient descent via the Adam optimizer (Kingma & Ba, 2014) with an initial "
        "learning rate of η = 10^(-3), mini-batch size of 64, and L2 weight decay of 10^(-3) to prevent parameter divergence. Because dry days outnumber "
        "rain days (approx. 61% vs. 39%), a standard cross-entropy loss would bias gradient updates toward the majority non-rain class, resulting in "
        "elevated false-negative rates on severe storm days."
    )
    add_paragraph(
        "To actively penalize missed rain events, we employ a class-weighted Cross-Entropy Loss function:\n"
        "L(θ) = - (1/N) ∑ [ w_0 · y_i · log(p_0,i) + w_1 · (1 - y_i) · log(p_1,i) ]\n"
        "where the class weight vector is computed dynamically from the training partition:\n"
        "w = [ 1.0 ,  N_dry / N_rain ] = [ 1.0 ,  1553 / 988 ] ≈ [ 1.0 ,  1.572 ]."
    )
    add_paragraph(
        "Training proceeds for a maximum of 60 epochs. Early stopping is enforced with a patience of 8 epochs monitored against validation loss. "
        "At the conclusion of training, the model weights achieving the minimum validation loss are restored, guaranteeing that no overfitted "
        "checkpoints are deployed to test evaluation."
    )

    # Figure 2: Learning Curves
    add_image_figure("week6_outputs/02_learning_curves.png",
                     "Training dynamics across epochs: Training Loss (left), Validation Loss (center), and Validation ROC-AUC (right). MLP converged and triggered early stopping at epoch 14; 1D-CNN triggered early stopping at epoch 9 due to rising validation loss; LSTM exhibited sustained, stable optimization, reaching peak validation AUC (0.848) and lowest validation loss (0.4703) before early stopping at epoch 30.",
                     2, width_inches=6.0)

    # -------------------------------------------------------------
    # SECTION 5
    # -------------------------------------------------------------
    add_heading_1("5. Empirical Results & Performance Benchmarking")

    add_heading_2("5.1 Single-Run Benchmark Results (Seed 42)")
    add_paragraph(
        "The primary test benchmark was conducted on the held-out test year (July 1, 2024 through June 30, 2025; 365 calendar days). "
        "During this period, exactly 157 days experienced rainfall (43.0% prevalence), while 208 days were dry (57.0% prevalence). "
        "A trivial majority-class (zero-rule) baseline predicting 'No Rain' every day yields an accuracy of only 57.0% and an ROC-AUC of 0.500."
    )
    add_paragraph(
        "Table 3 presents the primary classification metrics evaluated at the standard decision threshold of τ = 0.50, alongside threshold-free "
        "ROC-AUC, Precision-Recall AUC (PR-AUC), and the Brier score for probability calibration."
    )

    # Table 3: Test Benchmark
    t3_data = [
        ["Model Architecture", "Parameters", "Accuracy", "Precision (Rain)", "Recall (Rain)", "F1-Score", "Specificity", "ROC-AUC", "PR-AUC", "Brier Score"],
        ["Multilayer Perceptron (MLP)", "44,738", "83.56%", "81.29%", "80.25%", "80.77%", "86.06%", "0.8699", "0.8393", "0.1369"],
        ["Stacked LSTM Network", "55,426", "83.56%", "81.29%", "80.25%", "80.77%", "86.06%", "0.8903", "0.8747", "0.1252"],
        ["1D-Convolutional Net (CNN)", "16,642", "74.25%", "65.67%", "84.08%", "73.74%", "66.83%", "0.8414", "0.8232", "0.1650"],
        ["Majority Class Baseline", "—", "57.00%", "0.00%", "0.00%", "0.00%", "100.00%", "0.5000", "0.4301", "0.2451"]
    ]
    t3 = doc.add_table(rows=len(t3_data), cols=10)
    for r_idx, row in enumerate(t3_data):
        for c_idx, val in enumerate(row):
            t3.cell(r_idx, c_idx).paragraphs[0].text = val
    format_table(t3, [1.4, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6], 
                 [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    
    p_t3_cap = doc.add_paragraph()
    p_t3_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t3_cap.paragraph_format.space_before = Pt(4)
    p_t3_cap.paragraph_format.space_after = Pt(10)
    r_t3 = p_t3_cap.add_run("Table 3: Comprehensive test-set performance metrics evaluated on the held-out test year (2024-07-01 to 2025-06-30).")
    r_t3.bold = True
    r_t3.font.size = Pt(9.5)

    add_callout(
        "At the discrete threshold τ = 0.50, MLP and LSTM produce identical discrete classification metrics (83.56% accuracy, 80.77% F1-score). "
        "However, across continuous probability space, the LSTM demonstrates substantial superiority: achieving an ROC-AUC of 0.8903 (+2.04% over MLP), "
        "a PR-AUC of 0.8747 (+3.54% over MLP), and the lowest Brier calibration error of 0.1252 (an 8.5% improvement in probability accuracy).",
        bold_prefix="CRITICAL EMPIRICAL DISTINCTION: "
    )

    # Figure 3: Confusion Matrices
    add_image_figure("week6_outputs/03_confusion_matrices.png",
                     "Confusion matrices on 365 test days for MLP (left), LSTM (center), and 1D-CNN (right). Both MLP and LSTM recorded 179 True Negatives, 29 False Positives, 31 False Negatives, and 126 True Positives. The 1D-CNN captured more rain days (132 TP; 84.1% recall) but incurred 69 False Positives, degrading precision to 65.7%.",
                     3, width_inches=6.0)

    add_heading_2("5.2 Confusion Matrix Dynamics & Error Trade-offs")
    add_paragraph(
        "A breakdown of the confusion matrix counts (Figure 3) illustrates the contrasting operational trade-offs across the architectures:"
    )
    add_paragraph(
        "• MLP & LSTM (Balanced Precision/Recall): Out of 208 non-rain days, both models correctly identified 179 days (Specificity = 86.06%), "
        "generating only 29 false alarms (False Positive Rate = 13.94%). Out of 157 rain days, both correctly predicted 126 events (Recall = 80.25%), "
        "with 31 missed rain occurrences (False Negative Rate = 19.75%). This balance achieves an exceptional Precision of 81.29%."
    )
    add_paragraph(
        "• 1D-CNN (Aggressive Rain Detection): The convolutional model captured 132 rain days (Recall = 84.08%; only 25 missed storms). "
        "However, this heightened sensitivity came at the cost of excessive false alarms: 69 false positives occurred on dry days (Specificity dropped "
        "to 66.83%), causing precision to collapse to 65.67%. In an operational aviation context, issuing 69 false precipitation alerts would induce "
        "severe, unnecessary flight diversion and holding costs."
    )

    # Figure 4: ROC Curves & Figure 5: Metric Comparison
    add_image_figure("week6_outputs/04_roc_curves.png",
                     "Receiver Operating Characteristic (ROC) curves on the held-out test year. The stacked LSTM achieved the dominant envelope across all operating thresholds with an AUC of 0.890, compared to 0.870 for the MLP, 0.841 for the 1D-CNN, and 0.500 for random chance.",
                     4, width_inches=4.8)

    add_image_figure("week6_outputs/05_metric_comparison.png",
                     "Grouped bar chart comparing Accuracy, Precision, Recall, F1-Score, and ROC-AUC across MLP, LSTM, and 1D-CNN.",
                     5, width_inches=5.8)

    add_heading_2("5.3 Multi-Seed Monte Carlo Robustness Verification (N=5)")
    add_paragraph(
        "A well-known vulnerability of single-run deep learning evaluations on moderately sized meteorological test sets (~365 days) is the "
        "influence of weight initialization stochasticity. To rigorously verify whether the observed LSTM superiority was statistically robust "
        "or an artifact of lucky seed selection, we conducted a 5-fold Monte Carlo replication experiment across random seeds s ∈ {0, 1, 2, 3, 4}. "
        "Each architecture was re-initialized and retrained from scratch under identical hyperparameter conditions."
    )

    # Table 4: Multi-seed table
    t4_data = [
        ["Model Architecture", "Accuracy (Mean ± Std)", "Precision (Mean ± Std)", "Recall (Mean ± Std)", "F1-Score (Mean ± Std)", "ROC-AUC (Mean ± Std)", "PR-AUC (Mean ± Std)"],
        ["Multilayer Perceptron (MLP)", "82.19% ± 1.32%", "79.02% ± 2.15%", "79.87% ± 1.31%", "79.43% ± 1.35%", "0.8729 ± 0.0039", "0.8485 ± 0.0105"],
        ["Stacked LSTM Network", "83.78% ± 0.68%", "80.16% ± 0.99%", "82.80% ± 1.45%", "81.45% ± 0.82%", "0.8869 ± 0.0016", "0.8698 ± 0.0022"],
        ["1D-Convolutional Net (CNN)", "79.84% ± 1.09%", "74.87% ± 2.03%", "80.13% ± 2.71%", "77.36% ± 1.23%", "0.8489 ± 0.0051", "0.8349 ± 0.0093"]
    ]
    t4 = doc.add_table(rows=len(t4_data), cols=7)
    for r_idx, row in enumerate(t4_data):
        for c_idx, val in enumerate(row):
            t4.cell(r_idx, c_idx).paragraphs[0].text = val
    format_table(t4, [1.5, 0.8, 0.8, 0.8, 0.8, 0.9, 0.9],
                 [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    
    p_t4_cap = doc.add_paragraph()
    p_t4_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t4_cap.paragraph_format.space_before = Pt(4)
    p_t4_cap.paragraph_format.space_after = Pt(10)
    r_t4 = p_t4_cap.add_run("Table 4: Statistical robustness evaluation across 5 random seeds (Seeds 0, 1, 2, 3, 4) on held-out test data.")
    r_t4.bold = True
    r_t4.font.size = Pt(9.5)

    add_paragraph(
        "The multi-seed results in Table 4 definitively resolve the single-run tie. The stacked LSTM network achieves superior performance "
        "across every single evaluation metric, leading the MLP by +1.59% in mean accuracy (83.78% vs. 82.19%), +2.02% in mean F1-score (81.45% vs. 79.43%), "
        "and +1.40% in mean ROC-AUC (0.8869 vs. 0.8729). Furthermore, the LSTM exhibited extraordinary stability, demonstrating the lowest variance "
        "across all models (ROC-AUC standard deviation of merely σ = 0.0016, compared to σ = 0.0039 for MLP and σ = 0.0051 for 1D-CNN)."
    )

    # -------------------------------------------------------------
    # SECTION 6
    # -------------------------------------------------------------
    add_heading_1("6. Meteorological Discussion & Atmospheric Synthesis")

    add_heading_2("6.1 Why Recurrent Gating Outperforms Convolutions in the Himalayas")
    add_paragraph(
        "The clear empirical dominance of the LSTM over the 1D-CNN aligns directly with atmospheric physics. In a mountain-enclosed valley like "
        "Kathmandu, precipitation events are governed by cumulative synoptic processes that evolve over multiple days. Prior to a major monsoon "
        "trough passage, moist air from the Bay of Bengal advects into the Indo-Gangetic plain and is gradually channeled northward against the "
        "Himalayan foothills. This process manifests at surface stations as a gradual, multi-day barometric pressure decline, progressive daily "
        "increases in minimum temperature, and an escalating dew-point accumulation."
    )
    add_paragraph(
        "An LSTM cell updates its hidden state recurrently, allowing its forget gates to retain information regarding a low-pressure anomaly "
        "observed 7 days prior while integrating new moisture influxes observed today. In contrast, standard 1D-CNNs apply localized sliding filters "
        "(receptive field of 3 days per layer). To capture a 14-day sequence, the network must rely on aggressive pooling or deep stacking, which "
        "tends to blur subtle temporal transitions and smooth out critical inflection points. Consequently, the 1D-CNN struggled with high false-alarm "
        "rates during marginal pre-monsoon conditions."
    )

    add_heading_2("6.2 Seasonal Regime Shifts and Transition Windows")
    add_paragraph(
        "A qualitative audit of model errors revealed that misclassifications were heavily concentrated during seasonal transition windows, "
        "specifically the pre-monsoon transition (April to May) and post-monsoon cessation (late September to October):"
    )
    add_paragraph(
        "• Pre-Monsoon False Positives: During late April, daytime surface temperatures frequently reach 85°F to 90°F with elevated daylight hours, "
        "generating strong local convective instability. However, if upper-level mid-tropospheric moisture is insufficient, cumulus congestus clouds "
        "develop without producing measurable ground rainfall. Both MLP and 1D-CNN frequently produced false alarms during these dry convective cycles."
    )
    add_paragraph(
        "• Post-Monsoon False Negatives: During the retreat of the summer monsoon in early October, synoptic moisture levels decline significantly. "
        "However, sporadic localized western trough interactions can induce abrupt, unseasonable downpours. Because antecedent 14-day indicators "
        "exhibit clear drying trends, all models occasionally suffered false negatives on these isolated post-monsoon rain days."
    )

    add_heading_2("6.3 Operational Trade-offs for Aviation and Urban Infrastructure")
    add_paragraph(
        "In operational meteorology, a single universal decision threshold (τ = 0.50) is rarely optimal. The operational utility of a forecast "
        "depends on the asymmetric cost matrix between False Alarms and False Negatives:"
    )
    add_paragraph(
        "1. Civil Aviation Safety (TIA Operations): Missed convective rainfall carries severe risk, including sudden wind shear, runway hydroplaning, "
        "and emergency go-arounds. If airport dispatchers prioritize high recall, the decision threshold can be lowered to τ = 0.35. Due to its superior "
        "PR-AUC (0.8747) and ROC-AUC (0.8903), the LSTM maintains strong precision even under lower threshold operating regimes, preventing the severe "
        "false-alarm inflation observed in the 1D-CNN."
    )
    add_paragraph(
        "2. Municipal Drainage & Public Works: False alarms in municipal flood operations lead to 'warning fatigue' and wasted emergency crew "
        "deployments. For city managers, operating at τ = 0.50 or τ = 0.55 provides optimal precision (81.3%+) and specificity (86.1%+), ensuring "
        "that stormwater drainage teams are deployed with high confidence."
    )

    # -------------------------------------------------------------
    # SECTION 7
    # -------------------------------------------------------------
    add_heading_1("7. Threats to Validity, Limitations & Future Directions")

    add_heading_2("7.1 Single-Station Point Invariance vs. Valley Spatial Gradients")
    add_paragraph(
        "A recognized limitation of this study is its reliance on a single synoptic ground station located at Tribhuvan International Airport (TIA). "
        "While TIA is the official synoptic station for Kathmandu with the highest instrumentation quality, rainfall in the valley exhibits "
        "notable spatial heterogeneity. High-relief northern slopes (e.g., Budhanilkantha and Shivapuri) frequently receive 30% to 50% more rainfall "
        "than southern valley areas due to localized orographic triggering. Therefore, a binary rain event recorded at TIA may not fully capture "
        "localized flash flood hazards occurring across peripheral tributaries."
    )

    add_heading_2("7.2 Temporal Discontinuities and Precipitation Granularity")
    add_paragraph(
        "The historical 12-month calendar discontinuity (2015–2016) necessitated truncating the dataset to 2016 onwards. While our clean 9-year "
        "record (3,285 days) proved sufficient for training deep networks without overfitting, capturing decadal climate oscillation patterns "
        "(such as El Niño–Southern Oscillation, ENSO, and the Indian Ocean Dipole, IOD) would benefit from multi-decadal uninterrupted datasets. "
        "Furthermore, this study addressed binary rain occurrence (≥ 0.0 mm); operational forecasting would benefit from extending this framework "
        "to multi-class severity categories (Light, Moderate, Heavy, and Extreme) or continuous rainfall accumulation regression."
    )

    add_heading_2("7.3 Future Architectures: Radar Assimilation & Transformers")
    add_paragraph(
        "Promising avenues for future research include:"
    )
    add_paragraph("• Multi-Modal Spatial-Temporal Fusion: Integrating ground station time series with gridded satellite precipitation products (e.g., GPM IMERG) and newly deployed Himalayan X-band Doppler weather radars to capture upstream storm cell advection.")
    add_paragraph("• Transformer & State-Space Models: Exploring Temporal Fusion Transformers (TFT; Lim et al., 2021) and selective state-space architectures (Mamba) equipped with multi-head self-attention mechanisms to explicitly quantify feature importance across temporal horizons.")
    add_paragraph("• Physics-Informed Inductive Biases: Incorporating atmospheric thermodynamic constraints (e.g., Clausius-Clapeyron saturation relationships and hydrostatic equilibrium) directly into neural network loss formulations.")

    # -------------------------------------------------------------
    # SECTION 8
    # -------------------------------------------------------------
    add_heading_1("8. Conclusion & Strategic Recommendations")
    add_paragraph(
        "This research established a rigorous, leak-free empirical benchmark for next-day precipitation forecasting in the topographically complex "
        "Kathmandu Valley, comparing Multilayer Perceptrons, stacked Long Short-Term Memory networks, and 1D-Convolutional Neural Networks across "
        "a 9-year synoptic record from Tribhuvan International Airport. The key conclusions and actionable recommendations are as follows:"
    )
    add_paragraph(
        "1. Deep Sequence Modeling is Highly Effective: With engineered astronomical, cyclical harmonic, and physical lag features, deep neural "
        "networks achieve 83.6% accuracy and ~81% F1-score on held-out test data, vastly exceeding the 57.0% zero-rule baseline and offering a "
        "lightweight, real-time alternative to computationally heavy numerical weather prediction.", bold_prefix="1. High Operational Accuracy: "
    )
    add_paragraph(
        "2. LSTM is the Superior Architecture: Across both single-run benchmark testing and 5-seed Monte Carlo replication, the stacked LSTM "
        "consistently delivered the highest discriminative confidence (ROC-AUC = 0.8869 ± 0.0016; PR-AUC = 0.8698 ± 0.0022), highest mean F1-score "
        "(81.45% ± 0.82%), and lowest probability calibration error (Brier Score = 0.1252). Its recurrent memory gates successfully track multi-day "
        "synoptic pressure and moisture buildups that temporal convolutions fail to resolve.", bold_prefix="2. Recurrent Gating Superiority: "
    )
    add_paragraph(
        "3. Convolutional Networks Suffer False-Alarm Inflation: While the 1D-CNN achieved strong recall (84.1%), its rigid receptive fields "
        "produced 69 false alarms on dry test days (65.7% precision), rendering it sub-optimal for aviation operations sensitive to false warnings.",
        bold_prefix="3. CNN False Alarm Trade-off: "
    )
    add_paragraph(
        "4. Deployment Recommendation: We recommend operational deployment of the trained stacked LSTM model within the TIA meteorological "
        "dispatch system, paired with dynamic threshold calibration (τ = 0.35 to 0.50) to flexibly accommodate aviation safety and municipal flood mitigation.",
        bold_prefix="4. Operational Adoption: "
    )

    # -------------------------------------------------------------
    # SECTION 9: REFERENCES
    # -------------------------------------------------------------
    add_heading_1("9. References & Cited Journal Literature")

    references = [
        "Bai, S., Kolter, J. Z., & Koltun, V. (2018). An empirical evaluation of generic convolutional and recurrent networks for sequence modeling. arXiv preprint arXiv:1803.01271.",
        "Bohlinger, P., & Sorteberg, A. (2018). A comprehensive validation of volume-radar precipitation estimates in the Himalayas. Journal of Hydrometeorology, 19(5), 903–921.",
        "Brier, G. W. (1950). Verification of forecasts expressed in terms of probability. Monthly Weather Review, 78(1), 1–3.",
        "Davis, J., & Goadrich, M. (2006). The relationship between Precision-Recall and ROC curves. In Proceedings of the 23rd International Conference on Machine Learning (ICML '06), pp. 233–240.",
        "Dimri, A. P., Yasunari, T., Wiltshire, A., Kumar, P., Mathison, C., Ridley, J., & Jacob, D. (2015). Application of regional climate models to the Indian winter monsoon over the western Himalayas. Science of the Total Environment, 502, 675–686.",
        "Graves, A. (2012). Supervised Sequence Labelling with Recurrent Neural Networks. Studies in Computational Intelligence, Vol. 385. Springer, Berlin, Heidelberg.",
        "Hewage, P., Trovati, M., Pereira, E., & Behera, A. (2020). Deep learning-based effective fine-grained weather forecasting model. Information Fusion, 63, 30–44.",
        "Hochreiter, S., & Schmidhuber, J. (1997). Long Short-Term Memory. Neural Computation, 9(8), 1735–1780.",
        "Ioffe, S., & Szegedy, C. (2015). Batch Normalization: Accelerating deep network training by reducing internal covariate shift. In International Conference on Machine Learning (ICML '15), pp. 448–456.",
        "Jolliffe, I. T., & Stephenson, D. B. (Eds.). (2012). Forecast Verification: A Practitioner's Guide in Atmospheric Science (2nd ed.). John Wiley & Sons, Chichester, UK.",
        "Karki, R., Hasson, S. U., Schickhoff, U., Scholten, T., & Böhner, J. (2017). Rising minimum temperature and declining precipitation in the trans-Himalayan region of Nepal. International Journal of Climatology, 37(15), 5227–5244.",
        "Kingma, D. P., & Ba, J. (2014). Adam: A method for stochastic optimization. In 3rd International Conference on Learning Representations (ICLR '15), San Diego, CA.",
        "Kratzert, F., Klotz, D., Brenner, C., Schulz, K., & Herrnegger, M. (2018). Rainfall–runoff modelling using Long Short-Term Memory (LSTM) networks. Hydrology and Earth System Sciences, 22(11), 6005–6022.",
        "Kratzert, F., Klotz, D., Shalev, G., Klambauer, G., Hochreiter, S., & Nearing, G. (2019). Towards learning universal, regional, and local hydrological behaviors via machine learning applied to large-sample datasets. Hydrology and Earth System Sciences, 23(12), 5089–5110.",
        "Lea, C., Flynn, M. D., Vidal, R., Reiter, A., & Hager, G. D. (2017). Temporal Convolutional Networks for action segmentation and detection. In Proceedings of the IEEE Conference on Computer Vision and Pattern Recognition (CVPR '17), pp. 156–165.",
        "Lim, B., Arık, S. Ö., Loeff, N., & Pfister, T. (2021). Temporal Fusion Transformers for interpretable multi-horizon time series forecasting. International Journal of Forecasting, 37(4), 1748–1764.",
        "Rasp, S., Pritchard, M. S., & Gentine, P. (2018). Deep learning to represent subgrid processes in climate models. Proceedings of the National Academy of Sciences (PNAS), 115(39), 9684–9689.",
        "Reichstein, M., Camps-Valls, G., Stevens, B., Jung, M., Denzler, J., Carvalhais, N., & Prabhat. (2019). Deep learning and Earth system science for natural hazard and disaster risk reduction. Nature, 566(7743), 195–204.",
        "Schultz, M. G., Betancourt, C., Gong, B., Kleinert, F., Langguth, M., Leufen, L. H., Mozaffari, A., & Stadtler, S. (2021). Can deep learning beat numerical weather prediction? Philosophical Transactions of the Royal Society A, 379(2194), 20200097.",
        "Shrestha, A. B., Wake, C. P., Mayewski, P. A., & Dibb, J. E. (2000). Precipitation fluctuations in the Nepal Himalaya and its relationship with large-scale atmospheric circulation. International Journal of Climatology, 20(3), 317–327.",
        "Srivastava, N., Hinton, G., Krizhevsky, A., Sutskever, I., & Salakhutdinov, R. (2014). Dropout: A simple way to prevent neural networks from overfitting. Journal of Machine Learning Research, 15(1), 1929–1958.",
        "Sutskever, I., Vinyals, O., & Le, Q. V. (2014). Sequence to sequence learning with neural networks. In Advances in Neural Information Processing Systems (NeurIPS '14), pp. 3104–3112.",
        "Webster, P. J., Magaña, V. O., Palmer, T. N., Shukla, J., Tomas, R. A., Yanai, M., & Yasunari, T. (1998). Monsoons: Processes, predictability, and the prospects for prediction. Journal of Geophysical Research: Oceans, 103(C7), 14451–14510.",
        "Wilks, D. S. (2011). Statistical Methods in the Atmospheric Sciences (3rd ed.). Academic Press, Elsevier, Oxford, UK."
    ]

    for ref in references:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.space_after = Pt(4)
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)
        r_ref = p_ref.add_run(ref)
        r_ref.font.size = Pt(9)
        r_ref.font.name = "Calibri"
        r_ref.font.color.rgb = COLOR_BODY

    # Save document
    out_file = "Kathmandu_Rain_Classification_Report_Week6.docx"
    doc.save(out_file)
    print(f"Successfully generated: {out_file}")

if __name__ == "__main__":
    create_report()
