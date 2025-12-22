import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import numpy as np

# Page configuration
st.set_page_config(
    page_title="Q2 Training Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #f0f2f6;
        padding: 10px;
        border-radius: 5px;
    }
    /* Responsive header logo: scales with container, ~20% smaller */
    .header-logo img {
        width: 80%; /* 20% smaller than container */
        height: auto;
        max-width: 160px; /* cap size for large screens */
    }
    .center-title {
        color: #0a4a6e;
        font-weight: 700;
        letter-spacing: 0.3px;
        font-size: 2.5rem;
    }
    .center-subtitle {
        color: #d35400;
        font-weight: 600;
        font-size: 1.8rem;
        margin-top: 0.25rem;
    }
    </style>
    """, unsafe_allow_html=True)

# Load data
@st.cache_data
def load_data(file_path):
    programs_df = pd.read_excel(file_path, sheet_name='برامج الربع الثانى')
    trainees_df = pd.read_excel(file_path, sheet_name='بيانات المتدربين')
    registration_df = pd.read_excel(file_path, sheet_name='تسجيل المتدربين')
    return programs_df, trainees_df, registration_df

# Default file path - use relative path for deployment compatibility
import os
script_dir = Path(__file__).parent if '__file__' in globals() else Path.cwd()
default_file = script_dir / "second quarter.xlsx"

# File uploader in sidebar
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel File (optional)",
    type=['xlsx', 'xls'],
    help="Upload an updated Excel file with 3 sheets"
)

# Determine which file to use
if uploaded_file is not None:
    st.sidebar.success("✅ Using uploaded file")
    file_to_use = uploaded_file
else:
    st.sidebar.info("📂 Using default file: second quarter.xlsx")
    file_to_use = default_file

# Load the data
try:
    programs_df, trainees_df, registration_df = load_data(file_to_use)
    st.sidebar.success(f"✅ Data loaded successfully!")
    st.sidebar.metric("Last Updated", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"))
except Exception as e:
    st.error(f"❌ Error loading data: {str(e)}")
    st.stop()

# Top banner with logo and title
logo_path = script_dir / "assets" / "logo.png"
col_logo, col_title = st.columns([1, 5])
with col_logo:
    if logo_path.exists():
        import base64
        logo_bytes = logo_path.read_bytes()
        logo_b64 = base64.b64encode(logo_bytes).decode('utf-8')
        st.markdown(
            f'<div class="header-logo"><img src="data:image/png;base64,{logo_b64}" alt="Logo"/></div>',
            unsafe_allow_html=True
        )
    else:
        st.caption("Add logo.png to assets for header logo")

with col_title:
    st.markdown(
        """
        <div style="text-align:center; margin-top:0; margin-bottom:0.5rem;">
            <div class="center-title">Health Information Technology and Statistics Training Center</div>
            <div class="center-subtitle">Learn Today, Lead Tomorrow</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Title
st.title("📊 Second Quarter Training Dashboard")
st.markdown("---")

# Sidebar filters
st.sidebar.header("🔍 Filters")
sheet_view = st.sidebar.radio("Select View", ["Overview", "Programs", "Trainees", "Registrations", "Comparative Analysis"])

if sheet_view == "Overview":
    st.header("📈 Overview Statistics")
    
    # Calculate key metrics
    num_governorates = registration_df['المحافظة'].nunique()
    unique_programs = registration_df['البرنامج التدريبي'].nunique()
    total_courses = registration_df['البرنامج التدريبي ID'].nunique()
    actual_trainees = len(trainees_df)
    
    # Get target from programs sheet
    target_trainees = registration_df['عدد المستهدفين'].sum() if 'عدد المستهدفين' in registration_df.columns else 0
    
    # Key metrics row (target trainees metric removed per request)
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("عدد المحافظات", num_governorates)
    with col2:
        st.metric("عدد البرامج", unique_programs)
    with col3:
        st.metric("عدد الدورات", total_courses)
    with col4:
        st.metric("عدد المتدربين الفعلي", actual_trainees)
    
    st.markdown("---")
    
    # Gauge charts for KPIs
    col1, col2, col3 = st.columns(3)
    
    # Gauge 1: Programs Target Achievement
    with col1:
        target_programs = 13  # From image
        programs_delta = unique_programs - target_programs
        programs_pct = (unique_programs / target_programs * 100) if target_programs > 0 else 0
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=unique_programs,
            number={'suffix': '', 'valueformat': 'd'},
            delta={'reference': target_programs, 'valueformat': 'd', 'suffix': f' {programs_delta:+.0f}'},
            title={'text': f"عدد البرامج التدريبية من المستهدف<br><sub>Goal: {target_programs}</sub>"},
            gauge={
                'axis': {'range': [0, 15]},
                'bar': {'color': 'seagreen'},
                'steps': [
                    {'range': [0, 7], 'color': '#ffe6e6'},
                    {'range': [7, 11], 'color': '#fff4e6'},
                    {'range': [11, 15], 'color': '#e6f3ff'}
                ]
            }
        ))
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    # Gauge 2: Courses Target Achievement
    with col2:
        target_courses = 110  # From image
        courses_delta = total_courses - target_courses
        courses_pct = (total_courses / target_courses * 100) if target_courses > 0 else 0
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=total_courses,
            number={'suffix': '', 'valueformat': 'd'},
            delta={'reference': target_courses, 'valueformat': 'd', 'suffix': f' {courses_delta:+.0f}'},
            title={'text': f"عدد الدورات التدريبية الفعلي من المستهدف<br><sub>Goal: {target_courses}</sub>"},
            gauge={
                'axis': {'range': [0, 130]},
                'bar': {'color': 'royalblue'},
                'steps': [
                    {'range': [0, 55], 'color': '#ffe6e6'},
                    {'range': [55, 88], 'color': '#fff4e6'},
                    {'range': [88, 130], 'color': '#e6f3ff'}
                ]
            }
        ))
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    # Gauge 3: Trainees Fulfillment
    with col3:
        trainees_pct = (actual_trainees / target_trainees * 100) if target_trainees > 0 else 0
        trainees_delta = actual_trainees - target_trainees
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=trainees_pct,
            number={'suffix': '%', 'valueformat': '.2f'},
            delta={'reference': 100, 'valueformat': '.2f', 'suffix': '%'},
            title={'text': f"عدد المتدربين الفعلي من المستهدف<br><sub>الفعلي: {actual_trainees:,} | المستهدف: {int(target_trainees):,}</sub>"},
            gauge={
                'axis': {'range': [0, 100]},
                'bar': {'color': 'orange'},
                'threshold': {'line': {'color': 'green', 'width': 4}, 'value': 100},
                'steps': [
                    {'range': [0, 50], 'color': '#ffe6e6'},
                    {'range': [50, 80], 'color': '#fff4e6'},
                    {'range': [80, 100], 'color': '#e6f3ff'}
                ]
            }
        ))
        fig.update_layout(height=350)
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    
    # Exam Success Rates
    col1, col2, col3 = st.columns(3)
    
    # Calculate exam success rates
    initial_exam_pass = (pd.to_numeric(registration_df['نتيجة الإمتحان المبدئي'], errors='coerce') == 1).sum()
    final_exam_pass = (pd.to_numeric(registration_df['نتيجة الإمتحان النهائي'], errors='coerce') == 1).sum()
    total_registrations = len(registration_df)
    
    initial_success_rate = (initial_exam_pass / total_registrations * 100) if total_registrations > 0 else 0
    final_success_rate = (final_exam_pass / total_registrations * 100) if total_registrations > 0 else 0
    avg_success_rate = (initial_success_rate + final_success_rate) / 2
    
    with col1:
        st.metric("متوسط النجاح في الإمتحان", f"{avg_success_rate:.2f}%")
    with col2:
        st.metric("نسبة نجاح الإمتحان البدائي", f"{initial_success_rate:.2f}%")
    with col3:
        st.metric("نسبة نجاح الإمتحان النهائي", f"{final_success_rate:.2f}%")
    
    st.markdown("---")
    
    # Two column layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📚 Top Training Programs")
        program_counts = registration_df['البرنامج التدريبي'].value_counts().head(10)
        fig = px.bar(
            x=program_counts.values,
            y=program_counts.index,
            orientation='h',
            labels={'x': 'Registrations', 'y': 'Program'},
            title="Top 10 Training Programs"
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🏢 Training Locations")
        location_counts = programs_df['مكان التنفيذ'].value_counts()
        fig = px.pie(
            values=location_counts.values,
            names=location_counts.index,
            title="Programs by Location"
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # Full width charts
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📍 Training by Governorate")
        gov_counts = registration_df['مكان التدريب(محافظة)'].value_counts().head(10)
        fig = px.bar(
            x=gov_counts.index,
            y=gov_counts.values,
            labels={'x': 'Governorate', 'y': 'Registrations'},
            title="Top 10 Governorates"
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📅 Course Duration Distribution")
        duration_counts = pd.to_numeric(registration_df['عدد أيام الدورة'], errors='coerce').value_counts().sort_index()
        fig = px.bar(
            x=duration_counts.index,
            y=duration_counts.values,
            labels={'x': 'Course Duration (Days)', 'y': 'Registrations'},
            title="Distribution of Course Durations"
        )
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif sheet_view == "Programs":
    st.header("📋 Programs Analysis")
    
    # Metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Programs", len(programs_df))
    with col2:
        st.metric("Total Courses", len(programs_df))
    with col3:
        total_planned_capacity = len(programs_df) * 17.75
        st.metric("Planned Capacity", int(total_planned_capacity))
    with col4:
        unique_locations = programs_df['مكان التنفيذ'].nunique()
        st.metric("Training Locations", unique_locations)
    
    st.markdown("---")
    
    # Filters
    col1, col2 = st.columns(2)
    with col1:
        selected_location = st.multiselect(
            "Filter by Location",
            options=programs_df['مكان التنفيذ'].unique(),
            default=programs_df['مكان التنفيذ'].unique()
        )
    
    # Filter data
    filtered_programs = programs_df[
        programs_df['مكان التنفيذ'].isin(selected_location)
    ]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Programs by Location")
        location_dist = filtered_programs['مكان التنفيذ'].value_counts()
        fig = px.bar(
            x=location_dist.index,
            y=location_dist.values,
            labels={'x': 'Location', 'y': 'Number of Programs'},
            title="Program Distribution by Location"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("👥 Trainees per Course")
        trainees_per_course = filtered_programs['عدد المتدربين بالدورة الواحدة'].value_counts().sort_index()
        fig = px.bar(
            x=trainees_per_course.index,
            y=trainees_per_course.values,
            labels={'x': 'Trainees per Course', 'y': 'Count'},
            title="Distribution of Course Sizes"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Data table
    st.subheader("📑 Program Details")
    st.dataframe(
        filtered_programs,
        use_container_width=True,
        height=400
    )

elif sheet_view == "Trainees":
    st.header("👥 Trainees Analysis")
    
    # Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Trainees", len(trainees_df))
    with col2:
        unique_jobs = trainees_df['الوظيفة'].nunique()
        st.metric("Unique Job Titles", unique_jobs)
    with col3:
        unique_qualifications = trainees_df['المؤهل الدراسي'].nunique()
        st.metric("Education Levels", unique_qualifications)
    with col4:
        unique_workplaces = trainees_df['مكان العمل'].nunique()
        st.metric("Workplaces", unique_workplaces)
    with col5:
        unique_phones = trainees_df['رقم الموبايل'].nunique()
        st.metric("Unique Contacts", unique_phones)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("💼 Top Job Positions")
        job_counts = trainees_df['الوظيفة'].value_counts().head(10)
        fig = px.bar(
            x=job_counts.values,
            y=job_counts.index,
            orientation='h',
            labels={'x': 'Number of Trainees', 'y': 'Job Position'},
            title="Top 10 Job Positions"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🎓 Qualifications Distribution")
        qual_counts = trainees_df['المؤهل الدراسي'].value_counts().head(10)
        fig = px.pie(
            values=qual_counts.values,
            names=qual_counts.index,
            title="Trainees by Educational Qualification"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏢 Top Workplaces")
        workplace_counts = trainees_df['مكان العمل'].value_counts().head(10)
        fig = px.bar(
            x=workplace_counts.values,
            y=workplace_counts.index,
            orientation='h',
            labels={'x': 'Number of Trainees', 'y': 'Workplace'},
            title="Top 10 Workplaces"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📊 Trainees by Status")
        st.write(f"Total trainees in database: {len(trainees_df)}")
        st.write(f"Unique national IDs: {trainees_df['الرقم القومي'].nunique()}")
        st.write(f"Unique phone numbers: {trainees_df['رقم الموبايل'].nunique()}")
    
    # Data table
    st.subheader("📑 Trainee Details")
    display_cols = [' الاسم رباعي باللغة العربية', 'الوظيفة', 'مكان العمل', 'المؤهل الدراسي']
    st.dataframe(
        trainees_df[display_cols],
        use_container_width=True,
        height=400
    )

elif sheet_view == "Registrations":
    st.header("📝 Registration Analysis")
    
    # Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Registrations", len(registration_df))
    with col2:
        avg_attendance = pd.to_numeric(registration_df['Attendance'], errors='coerce').mean()
        st.metric("Avg Attendance", f"{avg_attendance:.1f}%")
    with col3:
        avg_duration = pd.to_numeric(registration_df['عدد أيام الدورة'], errors='coerce').mean()
        st.metric("Avg Course Days", f"{avg_duration:.1f}")
    with col4:
        unique_courses = registration_df['البرنامج التدريبي'].nunique()
        st.metric("Unique Programs", unique_courses)
    with col5:
        unique_locations = registration_df['مكان التدريب'].nunique()
        st.metric("Training Locations", unique_locations)
    
    st.markdown("---")
    
    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_programs = st.multiselect(
            "Filter by Program",
            options=registration_df['البرنامج التدريبي'].unique(),
            default=registration_df['البرنامج التدريبي'].unique()[:5]
        )
    with col2:
        selected_gov = st.multiselect(
            "Filter by Governorate",
            options=registration_df['مكان التدريب(محافظة)'].dropna().unique(),
            default=registration_df['مكان التدريب(محافظة)'].dropna().unique()[:5]
        )
    with col3:
        attendance_filter = st.slider(
            "Minimum Attendance %",
            min_value=0,
            max_value=100,
            value=0
        )
    
    # Filter data
    filtered_reg = registration_df[
        (registration_df['البرنامج التدريبي'].isin(selected_programs)) &
        (registration_df['مكان التدريب(محافظة)'].isin(selected_gov)) &
        (pd.to_numeric(registration_df['Attendance'], errors='coerce') >= attendance_filter)
    ]
    
    st.info(f"Showing {len(filtered_reg)} registrations based on filters")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Attendance Distribution")
        attendance_data = pd.to_numeric(filtered_reg['Attendance'], errors='coerce')
        fig = px.histogram(
            x=attendance_data,
            nbins=20,
            labels={'x': 'Attendance %', 'count': 'Number of Registrations'},
            title="Attendance Distribution"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📅 Exam Scores (Final)")
        final_exam = pd.to_numeric(filtered_reg['الامتحان النهائي'], errors='coerce')
        if final_exam.notna().sum() > 0:
            fig = px.histogram(
                x=final_exam,
                nbins=20,
                labels={'x': 'Final Exam Score', 'y': 'Count'},
                title="Final Exam Score Distribution"
            )
            st.plotly_chart(fig, use_container_width=True)
    
    # Data table
    st.subheader("📑 Registration Details")
    display_cols = ['البرنامج التدريبي', 'مكان التدريب(محافظة)', 'Attendance', 'الامتحان المبدئي', 'الامتحان النهائي']
    available_cols = [c for c in display_cols if c in filtered_reg.columns]
    st.dataframe(
        filtered_reg[available_cols],
        use_container_width=True,
        height=400
    )

else:  # Comparative Analysis
    st.header("🔄 Comparative Analysis")
    
    # Overall metrics
    st.subheader("📊 Overall Statistics")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Planned Capacity", int(len(programs_df) * 17.75))
    
    with col2:
        st.metric("Actual Registrations", len(registration_df))
    
    with col3:
        capacity = len(programs_df) * 17.75
        enrollment_rate = (len(registration_df) / capacity * 100) if capacity > 0 else 0
        st.metric("Enrollment Rate", f"{enrollment_rate:.1f}%")
    
    with col4:
        avg_attendance = pd.to_numeric(registration_df['Attendance'], errors='coerce').mean()
        st.metric("Avg Attendance", f"{avg_attendance:.1f}%")
    
    st.markdown("---")
    
    # Program comparison
    st.subheader("📚 Program-wise Comparison")
    
    # Aggregate by program
    plan_by_program = programs_df.groupby('البرنامج التدريبي').size().reset_index(name='planned_courses')
    actual_by_program = registration_df.groupby('البرنامج التدريبي').size().reset_index(name='actual_registrations')
    
    comparison = plan_by_program.merge(
        actual_by_program, 
        on='البرنامج التدريبي', 
        how='outer'
    ).fillna(0)
    
    comparison['planned_capacity'] = comparison['planned_courses'] * 17.75
    comparison['fulfillment_rate'] = (comparison['actual_registrations'] / comparison['planned_capacity'] * 100).round(1)
    comparison = comparison.sort_values('planned_courses', ascending=False).head(15)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Planned Capacity',
        x=comparison['البرنامج التدريبي'],
        y=comparison['planned_capacity'],
        marker_color='lightblue'
    ))
    fig.add_trace(go.Bar(
        name='Actual Registrations',
        x=comparison['البرنامج التدريبي'],
        y=comparison['actual_registrations'],
        marker_color='darkblue'
    ))
    
    fig.update_layout(
        title='Planned Capacity vs Actual Registrations (Top 15)',
        xaxis_title='Training Program',
        yaxis_title='Count',
        barmode='group',
        height=500
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Fulfillment rate
    st.subheader("✅ Fulfillment Rate by Program")
    fig = px.bar(
        comparison,
        x='البرنامج التدريبي',
        y='fulfillment_rate',
        labels={'fulfillment_rate': 'Fulfillment Rate (%)', 'البرنامج التدريبي': 'Program'},
        title='Program Fulfillment Rate (%)',
        color='fulfillment_rate',
        color_continuous_scale='RdYlGn'
    )
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Governorate comparison
    st.subheader("📍 Geographic Distribution")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Programs by Training Location**")
        location_dist = programs_df['مكان التنفيذ'].value_counts().head(10)
        fig = px.bar(
            x=location_dist.values,
            y=location_dist.index,
            orientation='h',
            labels={'x': 'Number of Programs', 'y': 'Location'},
            title="Top 10 Locations (Plan)"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.write("**Registrations by Governorate**")
        gov_dist = registration_df['مكان التدريب(محافظة)'].value_counts().head(10)
        fig = px.bar(
            x=gov_dist.values,
            y=gov_dist.index,
            orientation='h',
            labels={'x': 'Number of Registrations', 'y': 'Governorate'},
            title="Top 10 Governorates (Actual)"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Comparison table
    st.subheader("📋 Detailed Comparison Table")
    st.dataframe(
        comparison[['البرنامج التدريبي', 'planned_courses', 'planned_capacity', 'actual_registrations', 'fulfillment_rate']].rename(columns={
            'البرنامج التدريبي': 'Program',
            'planned_courses': 'Planned Courses',
            'planned_capacity': 'Planned Capacity',
            'actual_registrations': 'Registrations',
            'fulfillment_rate': 'Fulfillment Rate (%)'
        }),
        use_container_width=True,
        height=400
    )

# Footer
st.markdown("---")
st.markdown("**📊 Q2 Training Dashboard** | Data Source: second quarter.xlsx")
