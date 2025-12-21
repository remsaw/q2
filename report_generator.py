import pandas as pd
from pathlib import Path
import datetime as dt
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import arabic_reshaper
from bidi.algorithm import get_display

FILE_PATH = Path(r"c:\Q2\second quarter.xlsx")
OUTPUT_HTML = Path(r"c:\Q2\report.html")
OUTPUT_PDF = Path(r"c:\Q2\report.pdf")


def load_data(file_path: Path):
    programs_df = pd.read_excel(file_path, sheet_name='برامج الربع الثانى')
    trainees_df = pd.read_excel(file_path, sheet_name='بيانات المتدربين')
    registration_df = pd.read_excel(file_path, sheet_name='تسجيل المتدربين')
    return programs_df, trainees_df, registration_df


def safe_mean(series):
    try:
        return float(pd.to_numeric(series, errors='coerce').mean())
    except Exception:
        return float('nan')


def build_report(programs_df: pd.DataFrame, trainees_df: pd.DataFrame, registration_df: pd.DataFrame) -> str:
    total_courses = len(programs_df)
    total_trainees = len(trainees_df)
    total_registrations = len(registration_df)
    avg_attendance = safe_mean(registration_df.get('Attendance', pd.Series(dtype=float)))
    enrollment_rate = (total_registrations / (total_courses * 17.75) * 100) if total_courses > 0 else 0

    # Top programs and locations
    top_programs = registration_df['البرنامج التدريبي'].value_counts().head(10)
    location_counts = programs_df['مكان التنفيذ'].value_counts()

    # Governorate and durations
    gov_counts = registration_df['مكان التدريب(محافظة)'].value_counts().head(10)
    duration_counts = pd.to_numeric(registration_df['عدد أيام الدورة'], errors='coerce').value_counts().sort_index()

    # Program-wise comparison
    plan_by_program = programs_df.groupby('البرنامج التدريبي').size().reset_index(name='planned_courses')
    actual_by_program = registration_df.groupby('البرنامج التدريبي').size().reset_index(name='actual_registrations')
    comparison = plan_by_program.merge(actual_by_program, on='البرنامج التدريبي', how='outer').fillna(0)
    comparison['fulfillment_rate'] = (comparison['actual_registrations'] / (comparison['planned_courses'] * 17.75) * 100).fillna(0).round(1)
    comparison = comparison.sort_values('planned_courses', ascending=False)

    generated_at = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Basic HTML report
    html = [
        '<!DOCTYPE html>',
        '<html lang="en">',
        '<head>',
        '<meta charset="UTF-8"/>',
        '<meta name="viewport" content="width=device-width, initial-scale=1"/>',
        '<title>Q2 Training Full Report</title>',
        '<style>',
        'body{font-family:Segoe UI,Arial,Helvetica,sans-serif;margin:20px;color:#222}',
        'h1{margin-bottom:4px}',
        '.grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0}',
        '.card{background:#f7f9fc;border:1px solid #e5e9f2;border-radius:8px;padding:12px}',
        'table{width:100%;border-collapse:collapse;margin:10px 0}',
        'th,td{border:1px solid #ddd;padding:8px;text-align:left}',
        'th{background:#fafafa}',
        '.section{margin-top:24px}',
        '.muted{color:#666;font-size:12px}',
        '</style>',
        '</head>',
        '<body>',
        '<h1>📊 Q2 Training Full Report</h1>',
        f'<div class="muted">Generated at: {generated_at}</div>',
        '<div class="grid">',
        f'<div class="card"><div>Total Courses</div><h2>{total_courses}</h2></div>',
        f'<div class="card"><div>Total Trainees</div><h2>{total_trainees}</h2></div>',
        f'<div class="card"><div>Total Registrations</div><h2>{total_registrations}</h2></div>',
        f'<div class="card"><div>Enrollment Rate</div><h2>{enrollment_rate:.1f}%</h2></div>',
        '</div>',
        '<div class="section">',
        '<h2>Top Training Programs</h2>',
        '<table><thead><tr><th>Program</th><th>Registrations</th></tr></thead><tbody>',
    ]
    for name, count in top_programs.items():
        html.append(f'<tr><td>{name}</td><td>{int(count)}</td></tr>')
    html.extend(['</tbody></table>', '</div>'])

    # Training Locations
    html.extend(['<div class="section">', '<h2>Training Locations Distribution</h2>', '<table><thead><tr><th>Location</th><th>Courses</th></tr></thead><tbody>'])
    for name, count in location_counts.items():
        html.append(f'<tr><td>{name}</td><td>{int(count)}</td></tr>')
    html.extend(['</tbody></table>', '</div>'])

    # Governorates
    html.extend(['<div class="section">', '<h2>Top Governorates (Registrations)</h2>', '<table><thead><tr><th>Governorate</th><th>Registrations</th></tr></thead><tbody>'])
    for name, count in gov_counts.items():
        html.append(f'<tr><td>{name}</td><td>{int(count)}</td></tr>')
    html.extend(['</tbody></table>', '</div>'])

    # Durations
    html.extend(['<div class="section">', '<h2>Course Duration Distribution (Days)</h2>', '<table><thead><tr><th>Days</th><th>Registrations</th></tr></thead><tbody>'])
    for days, count in duration_counts.items():
        html.append(f'<tr><td>{int(days) if pd.notna(days) else ""}</td><td>{int(count)}</td></tr>')
    html.extend(['</tbody></table>', '</div>'])

    # Comparison
    html.extend(['<div class="section">', '<h2>Program-wise Planned vs Actual Registrations</h2>', '<table><thead><tr><th>Program</th><th>Planned Courses</th><th>Registrations</th><th>Fulfillment %</th></tr></thead><tbody>'])
    for _, row in comparison.iterrows():
        prog = row['البرنامج التدريبي']
        planned = int(row['planned_courses'])
        actual = int(row['actual_registrations'])
        rate = float(row['fulfillment_rate'])
        html.append(f'<tr><td>{prog}</td><td>{planned}</td><td>{actual}</td><td>{rate:.1f}</td></tr>')
    html.extend(['</tbody></table>', '</div>'])

    html.extend(['<hr/>', '<div class="muted">Data Source: second quarter.xlsx</div>', '</body>', '</html>'])
    return "\n".join(html)


def build_pdf(programs_df: pd.DataFrame, trainees_df: pd.DataFrame, registration_df: pd.DataFrame):
    def shape(text: str) -> str:
        try:
            if not isinstance(text, str):
                text = str(text)
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        except Exception:
            return str(text)

    total_courses = len(programs_df)
    total_trainees = len(trainees_df)
    total_registrations = len(registration_df)
    avg_attendance = safe_mean(registration_df.get('Attendance', pd.Series(dtype=float)))
    enrollment_rate = (total_registrations / (total_courses * 17.75) * 100) if total_courses > 0 else 0

    # Prepare summaries
    top_programs = registration_df['البرنامج التدريبي'].value_counts().head(10)
    location_counts = programs_df['مكان التنفيذ'].value_counts()
    gov_counts = registration_df['مكان التدريب(محافظة)'].value_counts().head(10)

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # Register a Unicode-capable font (Windows Tahoma supports Arabic)
    font_path = Path('C:/Windows/Fonts/Tahoma.ttf')
    bold_font_path = Path('C:/Windows/Fonts/Tahomabd.ttf')
    if font_path.exists():
        pdf.add_font('Tahoma', '', str(font_path))
        if bold_font_path.exists():
            pdf.add_font('Tahoma', 'B', str(bold_font_path))
        base_font = 'Tahoma'
    else:
        # Fallback to core font (may not render Arabic correctly)
        base_font = 'Arial'

    # Title
    pdf.set_font(base_font, 'B', 16)
    pdf.cell(0, 10, shape('Q2 Training Full Report'), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(base_font, '', 10)
    pdf.cell(0, 6, shape(f"Generated at: {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)

    # Key metrics
    pdf.set_font(base_font, 'B', 12)
    pdf.cell(0, 8, shape('Key Metrics'), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(base_font, '', 11)
    metrics = [
        ("Total Courses", str(total_courses)),
        ("Total Trainees", str(total_trainees)),
        ("Total Registrations", str(total_registrations)),
        ("Enrollment Rate", f"{enrollment_rate:.1f}%"),
        ("Avg Attendance", f"{avg_attendance:.1f}%" if avg_attendance == avg_attendance else "N/A"),
    ]
    for k, v in metrics:
        pdf.cell(0, 6, shape(f"- {k}: {v}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Section helper
    def section(title: str):
        pdf.ln(3)
        pdf.set_font(base_font, 'B', 12)
        pdf.cell(0, 8, title, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font(base_font, '', 11)

    # Top programs
    section('Top Training Programs')
    for name, count in top_programs.items():
        pdf.cell(0, 6, shape(f"{name}: {int(count)}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Training Locations
    section('Training Locations Distribution')
    for name, count in location_counts.items():
        pdf.cell(0, 6, shape(f"{name}: {int(count)} courses"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Governorates
    section('Top Governorates (Registrations)')
    for name, count in gov_counts.items():
        pdf.cell(0, 6, shape(f"{name}: {int(count)}"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Plan vs Actual
    section('Program-wise Planned vs Actual Registrations (Top)')
    plan_by_program = programs_df.groupby('البرنامج التدريبي').size().reset_index(name='planned_courses')
    actual_by_program = registration_df.groupby('البرنامج التدريبي').size().reset_index(name='actual_registrations')
    comparison = plan_by_program.merge(actual_by_program, on='البرنامج التدريبي', how='outer').fillna(0)
    comparison['fulfillment_rate'] = (comparison['actual_registrations'] / (comparison['planned_courses'] * 17.75) * 100).fillna(0).round(1)
    comparison = comparison.sort_values('planned_courses', ascending=False).head(15)
    for _, row in comparison.iterrows():
        prog = str(row['البرنامج التدريبي'])
        planned = int(row['planned_courses'])
        actual = int(row['actual_registrations'])
        rate = float(row['fulfillment_rate'])
        pdf.cell(0, 6, shape(f"{prog} — Planned: {planned}, Registrations: {actual}, Fulfillment: {rate:.1f}%"), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.ln(4)
    pdf.set_font(base_font, '', 9)
    pdf.cell(0, 6, shape('Data Source: second quarter.xlsx'), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    try:
        pdf.output(str(OUTPUT_PDF))
    except PermissionError:
        alt = OUTPUT_PDF.with_name('report_ar.pdf')
        pdf.output(str(alt))
        print(f"PDF was open; saved as: {alt}")


def main():
    programs_df, trainees_df, registration_df = load_data(FILE_PATH)
    html = build_report(programs_df, trainees_df, registration_df)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print(f"Report written to: {OUTPUT_HTML}")
    build_pdf(programs_df, trainees_df, registration_df)
    print(f"PDF written to: {OUTPUT_PDF}")


if __name__ == '__main__':
    main()
