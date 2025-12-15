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
    </style>
    """, unsafe_allow_html=True)

# Load data
@st.cache_data
def load_data(file_path):
    plan_df = pd.read_excel(file_path, sheet_name='plan')
    trainee_df = pd.read_excel(file_path, sheet_name='trainne')
    return plan_df, trainee_df

# Default file path (if the file is bundled with the app repo)
default_file = (Path(__file__).parent / "second quarter.xlsx")

# File uploader in sidebar
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader(
    "Upload Excel File (optional)",
    type=['xlsx', 'xls'],
    help="Upload an updated Excel file with 'plan' and 'trainne' sheets"
)

# Determine which file to use
if uploaded_file is not None:
    st.sidebar.success("✅ Using uploaded file")
    file_to_use = uploaded_file
elif default_file.exists():
    st.sidebar.info("📂 Using bundled file: second quarter.xlsx")
    file_to_use = default_file
else:
    file_to_use = None

# Load the data
try:
    if file_to_use is None:
        st.warning("يرجى رفع ملف Excel يحتوي على sheet بإسم plan و trainne؛ لا يوجد ملف افتراضي على الخادم.")
        st.stop()
    plan_df, trainee_df = load_data(file_to_use)
    st.sidebar.success("✅ Data loaded successfully!")
    st.sidebar.metric("Last Updated", pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S"))
except Exception as e:
    st.error(f"❌ Error loading data: {str(e)}")
    st.stop()

# Title
st.title("📊 Second Quarter Training Dashboard")
st.markdown("---")

# Sidebar filters
st.sidebar.header("🔍 Filters")
sheet_view = st.sidebar.radio("Select View", ["Overview", "Training Plan", "Trainee Details", "Comparative Analysis"])

if sheet_view == "Overview":
    st.header("📈 Overview Statistics")
    
    # Calculate performance metrics
    passing_threshold = 60
    trainees_passed = (trainee_df['Attendance'] >= passing_threshold).sum()
    trainees_failed = (trainee_df['Attendance'] < passing_threshold).sum()
    total_trainees = len(trainee_df)
    success_rate = (trainees_passed / total_trainees * 100) if total_trainees > 0 else 0
    failure_rate = (trainees_failed / total_trainees * 100) if total_trainees > 0 else 0
    avg_score = trainee_df['Attendance'].mean()
    median_score = trainee_df['Attendance'].median()
    std_dev = trainee_df['Attendance'].std()
    
    # Key metrics - Row 1
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Courses", len(plan_df))
    with col2:
        st.metric("Total Trainees", len(trainee_df))
    with col3:
        st.metric("Target Participants", plan_df['عدد المستهدفين'].sum())
    with col4:
        st.metric("Avg Attendance", f"{avg_score:.1f}%")
    
    # Key metrics - Row 2 (Performance Indicators)
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("معدل النجاح (Success Rate)", f"{success_rate:.1f}%", delta=f"{trainees_passed} ناجح")
    with col2:
        st.metric("معدل الرسوب (Failure Rate)", f"{failure_rate:.1f}%", delta=f"{trainees_failed} راسب")
    with col3:
        st.metric("المتوسط العام (Mean Score)", f"{avg_score:.2f}", delta="نقطة من 100")
    with col4:
        st.metric("الانحراف المعياري (Std Dev)", f"{std_dev:.2f}", delta="تشتت الدرجات")
    
    st.markdown("---")
    
    # Performance by Governorate and Program
    col1, col2 = st.columns(2)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏆 أعلى 5 محافظات أداءً (Top 5 Governorates)")
        gov_performance = trainee_df.groupby('المحافظة').apply(
            lambda x: {
                'success_count': (x['Attendance'] >= 60).sum(),
                'total': len(x),
                'success_rate': (x['Attendance'] >= 60).sum() / len(x) * 100,
                'avg_score': x['Attendance'].mean()
            }
        ).reset_index()
        gov_performance.columns = ['المحافظة', 'metrics']
        gov_performance = gov_performance.dropna()
        gov_performance['success_rate'] = gov_performance['metrics'].apply(lambda x: x['success_rate'])
        gov_performance = gov_performance.sort_values('success_rate', ascending=False).head(5)
        
        fig = px.bar(
            gov_performance,
            x='success_rate',
            y='المحافظة',
            orientation='h',
            labels={'success_rate': 'معدل النجاح (%)', 'المحافظة': 'المحافظة'},
            title="أعلى 5 محافظات أداءً",
            color='success_rate',
            color_continuous_scale='Greens'
        )
        fig.update_traces(text=gov_performance['success_rate'].round(2), textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🎯 أعلى 5 برامج أداءً (Top 5 Programs)")
        program_performance = trainee_df.groupby('البرنامج التدريبي').apply(
            lambda x: {
                'success_count': (x['Attendance'] >= 60).sum(),
                'total': len(x),
                'success_rate': (x['Attendance'] >= 60).sum() / len(x) * 100,
                'avg_score': x['Attendance'].mean()
            }
        ).reset_index()
        program_performance.columns = ['البرنامج التدريبي', 'metrics']
        program_performance = program_performance.dropna()
        program_performance['success_rate'] = program_performance['metrics'].apply(lambda x: x['success_rate'])
        program_performance = program_performance.sort_values('success_rate', ascending=False).head(5)
        
        fig = px.bar(
            program_performance,
            x='success_rate',
            y='البرنامج التدريبي',
            orientation='h',
            labels={'success_rate': 'معدل النجاح (%)', 'البرنامج التدريبي': 'البرنامج'},
            title="أعلى 5 برامج أداءً",
            color='success_rate',
            color_continuous_scale='Greens'
        )
        fig.update_traces(text=program_performance['success_rate'].round(2), textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    
    # Two column layout
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📚 Top Training Programs")
        course_counts = trainee_df['البرنامج التدريبي'].value_counts().head(10)
        fig = px.bar(
            x=course_counts.values,
            y=course_counts.index,
            orientation='h',
            labels={'x': 'Number of Trainees', 'y': 'Training Program'},
            title="Top 10 Training Programs by Enrollment"
        )
        fig.update_traces(text=course_counts.values, textposition='outside')
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🏢 Departments Distribution")
        dept_counts = plan_df['الإدارة'].value_counts()
        fig = px.pie(
            values=dept_counts.values,
            names=dept_counts.index,
            title="Training Courses by Department"
        )
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    # Full width charts
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📍 Training by Governorate")
        gov_counts = trainee_df['المحافظة'].value_counts().head(10)
        fig = px.bar(
            x=gov_counts.index,
            y=gov_counts.values,
            labels={'x': 'Governorate', 'y': 'Number of Trainees'},
            title="Top 10 Governorates by Trainee Count"
        )
        fig.update_traces(text=gov_counts.values, textposition='outside')
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📅 Course Duration Distribution")
        duration_counts = trainee_df['عدد أيام الدورة'].value_counts().sort_index()
        fig = px.bar(
            x=duration_counts.index,
            y=duration_counts.values,
            labels={'x': 'Course Duration (Days)', 'y': 'Number of Trainees'},
            title="Distribution of Course Durations"
        )
        fig.update_traces(text=duration_counts.values, textposition='outside')
        fig.update_layout(height=400, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

elif sheet_view == "Training Plan":
    st.header("📋 Training Plan Analysis")
    
    # Metrics
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Courses", len(plan_df))
    with col2:
        st.metric("Total Target", plan_df['عدد المستهدفين'].sum())
    with col3:
        st.metric("Avg Target/Course", f"{plan_df['عدد المستهدفين'].mean():.1f}")
    with col4:
        unique_locations = plan_df['مكان التدريب'].nunique()
        st.metric("Training Locations", unique_locations)
    
    st.markdown("---")
    
    # Filters
    col1, col2 = st.columns(2)
    with col1:
        selected_dept = st.multiselect(
            "Filter by Department",
            options=plan_df['الإدارة'].unique(),
            default=plan_df['الإدارة'].unique()
        )
    with col2:
        selected_gov = st.multiselect(
            "Filter by Governorate",
            options=plan_df['المحافظة'].unique(),
            default=plan_df['المحافظة'].unique()
        )
    
    # Filter data
    filtered_plan = plan_df[
        (plan_df['الإدارة'].isin(selected_dept)) &
        (plan_df['المحافظة'].isin(selected_gov))
    ]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Courses by Department")
        dept_dist = filtered_plan['الإدارة'].value_counts()
        fig = px.bar(
            x=dept_dist.index,
            y=dept_dist.values,
            labels={'x': 'Department', 'y': 'Number of Courses'},
            title="Course Distribution by Department"
        )
        fig.update_traces(text=dept_dist.values, textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🎯 Target Participants by Program")
        program_targets = filtered_plan.groupby('البرنامج التدريبي')['عدد المستهدفين'].sum().sort_values(ascending=False).head(10)
        fig = px.bar(
            x=program_targets.values,
            y=program_targets.index,
            orientation='h',
            labels={'x': 'Target Participants', 'y': 'Training Program'},
            title="Top Programs by Target Participants"
        )
        fig.update_traces(text=program_targets.values, textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    # Timeline
    st.subheader("📅 Training Schedule Timeline")
    timeline_data = filtered_plan[['البرنامج التدريبي', 'بداية الدورة', 'نهاية الدورة']].copy()
    timeline_data = timeline_data.dropna()
    
    if len(timeline_data) > 0:
        fig = px.timeline(
            timeline_data,
            x_start='بداية الدورة',
            x_end='نهاية الدورة',
            y='البرنامج التدريبي',
            title="Course Timeline"
        )
        fig.update_layout(height=600)
        st.plotly_chart(fig, use_container_width=True)
    
    # Data table
    st.subheader("📑 Filtered Course Data")
    st.dataframe(
        filtered_plan[['م', 'الإدارة', 'البرنامج التدريبي', 'المحافظة', 'بداية الدورة', 'نهاية الدورة', 'عدد المستهدفين']],
        use_container_width=True,
        height=400
    )

elif sheet_view == "Trainee Details":
    st.header("👥 Trainee Analysis")
    
    # Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Trainees", len(trainee_df))
    with col2:
        avg_attendance = trainee_df['Attendance'].mean()
        st.metric("Avg Attendance", f"{avg_attendance:.1f}%")
    with col3:
        avg_duration = trainee_df['عدد أيام الدورة'].mean()
        st.metric("Avg Course Days", f"{avg_duration:.1f}")
    with col4:
        unique_courses = trainee_df['البرنامج التدريبي'].nunique()
        st.metric("Unique Courses", unique_courses)
    with col5:
        unique_locations = trainee_df['مكان التدريب'].nunique()
        st.metric("Training Locations", unique_locations)
    
    st.markdown("---")
    
    # Filters
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_course = st.multiselect(
            "Filter by Course",
            options=trainee_df['البرنامج التدريبي'].unique(),
            default=trainee_df['البرنامج التدريبي'].unique()[:5]
        )
    with col2:
        selected_gov_trainee = st.multiselect(
            "Filter by Governorate",
            options=trainee_df['المحافظة'].dropna().unique(),
            default=trainee_df['المحافظة'].dropna().unique()[:5]
        )
    with col3:
        attendance_filter = st.slider(
            "Minimum Attendance %",
            min_value=0,
            max_value=100,
            value=0
        )
    
    # Filter data
    filtered_trainee = trainee_df[
        (trainee_df['البرنامج التدريبي'].isin(selected_course)) &
        (trainee_df['المحافظة'].isin(selected_gov_trainee)) &
        (trainee_df['Attendance'] >= attendance_filter)
    ]
    
    st.info(f"Showing {len(filtered_trainee)} trainees based on filters")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Attendance Distribution")
        fig = px.histogram(
            filtered_trainee,
            x='Attendance',
            nbins=20,
            labels={'Attendance': 'Attendance %', 'count': 'Number of Trainees'},
            title="Trainee Attendance Distribution"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🎓 Qualifications Distribution")
        qual_counts = filtered_trainee['المؤهل الدراسي'].value_counts().head(10)
        fig = px.pie(
            values=qual_counts.values,
            names=qual_counts.index,
            title="Trainees by Educational Qualification"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("💼 Top Job Positions")
        job_counts = filtered_trainee['الوظيفة'].value_counts().head(10)
        fig = px.bar(
            x=job_counts.values,
            y=job_counts.index,
            orientation='h',
            labels={'x': 'Number of Trainees', 'y': 'Job Position'},
            title="Top 10 Job Positions"
        )
        fig.update_traces(text=job_counts.values, textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("🏢 Top Workplaces")
        workplace_counts = filtered_trainee['مكان العمل'].value_counts().head(10)
        fig = px.bar(
            x=workplace_counts.values,
            y=workplace_counts.index,
            orientation='h',
            labels={'x': 'Number of Trainees', 'y': 'Workplace'},
            title="Top 10 Workplaces"
        )
        fig.update_traces(text=workplace_counts.values, textposition='outside')
        st.plotly_chart(fig, use_container_width=True)
    
    # Exam scores analysis (if available)
    st.subheader("📝 Exam Performance")
    
    # Try to convert exam scores to numeric
    try:
        # Normalize possible percent strings like "72.3%" and 0-1 scaled values
        def _to_number_percent_aware(series: pd.Series) -> pd.Series:
            if series is None:
                return pd.Series(dtype='float64')
            s = series.astype(str).str.strip()
            # Remove percent sign and normalize decimal separator
            s = s.str.replace('%', '', regex=False).str.replace(',', '.', regex=False)
            nums = pd.to_numeric(s, errors='coerce')
            # If most values look between 0 and 1, scale to 0-100
            non_na = nums.dropna()
            if len(non_na) > 0 and (non_na.between(0, 1).mean() > 0.8):
                nums = nums * 100
            return nums
        # Initial exam (always percent-like)
        trainee_df['الامتحان المبدئي_numeric'] = _to_number_percent_aware(trainee_df.get('الامتحان المبدئي'))
        # Final exam may be in 'Score' or 'الامتحان النهائي' — choose the one with data
        score_series = _to_number_percent_aware(trainee_df['Score']) if 'Score' in trainee_df.columns else pd.Series(dtype='float64')
        alt_final_series = _to_number_percent_aware(trainee_df['الامتحان النهائي']) if 'الامتحان النهائي' in trainee_df.columns else pd.Series(dtype='float64')
        score_non_na = score_series.notna().sum() if not score_series.empty else 0
        alt_non_na = alt_final_series.notna().sum() if not alt_final_series.empty else 0
        threshold = max(5, int(0.05 * len(trainee_df))) if len(trainee_df) > 0 else 0
        if score_non_na >= threshold:
            trainee_df['final_exam_numeric'] = score_series
            final_source_used = 'Score'
        elif alt_non_na > 0:
            trainee_df['final_exam_numeric'] = alt_final_series
            final_source_used = 'الامتحان النهائي'
        else:
            trainee_df['final_exam_numeric'] = pd.Series([None] * len(trainee_df))
            final_source_used = 'غير متاح'
        
        # Calculate success rate by program (60 is passing), over valid rows only
        def _success_rates(g: pd.DataFrame):
            init_valid = g['الامتحان المبدئي_numeric'].notna().sum()
            final_valid = g['final_exam_numeric'].notna().sum()
            init_rate = ((g['الامتحان المبدئي_numeric'] >= 60).sum() / init_valid * 100) if init_valid > 0 else 0
            final_rate = ((g['final_exam_numeric'] >= 60).sum() / final_valid * 100) if final_valid > 0 else 0
            return pd.Series({'initial_success_rate': init_rate, 'final_success_rate': final_rate})

        program_exam_data = trainee_df.groupby('البرنامج التدريبي').apply(_success_rates).reset_index()
        
        program_exam_data.columns = ['البرنامج التدريبي', 'metrics']
        program_exam_data['initial_success_rate'] = program_exam_data['metrics'].apply(lambda x: x['initial_success_rate'])
        program_exam_data['final_success_rate'] = program_exam_data['metrics'].apply(lambda x: x['final_success_rate'])
        program_exam_data = program_exam_data.drop('metrics', axis=1)
        program_exam_data = program_exam_data.sort_values('final_success_rate', ascending=False)
        
        if len(program_exam_data) > 0 and (
            program_exam_data['initial_success_rate'].notna().any() or program_exam_data['final_success_rate'].notna().any()
        ) and (
            program_exam_data['initial_success_rate'].sum() > 0 or program_exam_data['final_success_rate'].sum() > 0
        ):
            st.subheader("📊 متوسط نسبة النجاح / البرنامج التدريبي (Average Success Rate by Program)")
            st.caption(f"مصدر الامتحان النهائي المستخدم: {final_source_used}")
            
            fig = go.Figure()
            fig.add_trace(go.Bar(
                name='الامتحان المبدئي (Initial Exam)',
                x=program_exam_data['البرنامج التدريبي'],
                y=program_exam_data['initial_success_rate'],
                marker_color='crimson',
                text=program_exam_data['initial_success_rate'].round(2),
                textposition='outside'
            ))
            fig.add_trace(go.Bar(
                name='الامتحان النهائي (Final Exam)',
                x=program_exam_data['البرنامج التدريبي'],
                y=program_exam_data['final_success_rate'],
                marker_color='navy',
                text=program_exam_data['final_success_rate'].round(2),
                textposition='outside'
            ))
            
            fig.update_layout(
                title='متوسط نسبة النجاح / البرنامج التدريبي',
                xaxis_title='البرنامج التدريبي',
                yaxis_title='نسبة النجاح (%)',
                barmode='group',
                height=500,
                xaxis_tickangle=-45,
                hovermode='x unified'
            )
            st.plotly_chart(fig, use_container_width=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if trainee_df['الامتحان المبدئي_numeric'].notna().sum() > 0:
                fig = px.histogram(
                    trainee_df.dropna(subset=['الامتحان المبدئي_numeric']),
                    x='الامتحان المبدئي_numeric',
                    nbins=20,
                    labels={'الامتحان المبدئي_numeric': 'Initial Exam (%)'},
                    title="Initial Exam % Distribution"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            if trainee_df['final_exam_numeric'].notna().sum() > 0:
                fig = px.histogram(
                    trainee_df.dropna(subset=['final_exam_numeric']),
                    x='final_exam_numeric',
                    nbins=20,
                    labels={'final_exam_numeric': 'Final Exam (Score)'},
                    title="Final Exam Score Distribution"
                )
                st.plotly_chart(fig, use_container_width=True)
    except:
        st.info("Exam scores are not in numeric format for analysis")
    
    # Data table
    st.subheader("📑 Filtered Trainee Data")
    display_cols = [' الاسم رباعي باللغة العربية', 'البرنامج التدريبي', 'المحافظة', 
                   'الوظيفة', 'المؤهل الدراسي', 'Attendance', 'عدد أيام الدورة']
    st.dataframe(
        filtered_trainee[display_cols],
        use_container_width=True,
        height=400
    )

else:  # Comparative Analysis
    st.header("🔄 Comparative Analysis")
    
    # Plan vs Actual
    st.subheader("📊 Plan vs Actual Enrollment")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Planned Target", plan_df['عدد المستهدفين'].sum())
        st.metric("Planned Courses", len(plan_df))
    
    with col2:
        st.metric("Actual Trainees", len(trainee_df))
        enrollment_rate = (len(trainee_df) / plan_df['عدد المستهدفين'].sum()) * 100
        st.metric("Enrollment Rate", f"{enrollment_rate:.1f}%")
    
    st.markdown("---")
    
    # Program comparison
    st.subheader("📚 Program-wise Comparison")
    
    # Aggregate by program
    plan_by_program = plan_df.groupby('البرنامج التدريبي')['عدد المستهدفين'].sum().reset_index()
    actual_by_program = trainee_df.groupby('البرنامج التدريبي').size().reset_index(name='actual_count')
    
    comparison = plan_by_program.merge(
        actual_by_program, 
        on='البرنامج التدريبي', 
        how='outer'
    ).fillna(0)
    
    comparison['fulfillment_rate'] = (comparison['actual_count'] / comparison['عدد المستهدفين'] * 100).round(1)
    comparison = comparison.sort_values('عدد المستهدفين', ascending=False).head(15)
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Target',
        x=comparison['البرنامج التدريبي'],
        y=comparison['عدد المستهدفين'],
        marker_color='lightblue',
        text=comparison['عدد المستهدفين'],
        textposition='outside'
    ))
    fig.add_trace(go.Bar(
        name='Actual',
        x=comparison['البرنامج التدريبي'],
        y=comparison['actual_count'],
        marker_color='darkblue',
        text=comparison['actual_count'],
        textposition='outside'
    ))
    
    fig.update_layout(
        title='Target vs Actual Enrollment by Program (Top 15)',
        xaxis_title='Training Program',
        yaxis_title='Number of Participants',
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
        labels={'fulfillment_rate': 'Fulfillment Rate (%)', 'البرنامج التدريبي': 'Training Program'},
        title='Program Fulfillment Rate (%)',
        color='fulfillment_rate',
        color_continuous_scale='RdYlGn'
    )
    fig.update_traces(text=comparison['fulfillment_rate'], textposition='outside')
    fig.update_layout(height=400)
    st.plotly_chart(fig, use_container_width=True)
    
    # Governorate comparison - Plan vs Actual
    st.subheader("📍 Planned vs Actual by Governorate")
    
    # Prepare data for governorate comparison
    plan_by_gov = plan_df.groupby('المحافظة')['عدد المستهدفين'].sum().reset_index()
    actual_by_gov = trainee_df['المحافظة'].value_counts().reset_index()
    actual_by_gov.columns = ['المحافظة', 'actual_count']
    
    gov_comparison = plan_by_gov.merge(actual_by_gov, on='المحافظة', how='outer').fillna(0)
    gov_comparison = gov_comparison.sort_values('عدد المستهدفين', ascending=False)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**عدد الدورات / محافظة** (Number of Courses / Governorate)")
        plan_courses_gov = plan_df['المحافظة'].value_counts().reset_index()
        plan_courses_gov.columns = ['المحافظة', 'planned_courses']
        actual_courses_gov = trainee_df.groupby('المحافظة').size().reset_index(name='actual_courses')
        
        courses_gov_comp = plan_courses_gov.merge(actual_courses_gov, on='المحافظة', how='outer').fillna(0)
        courses_gov_comp = courses_gov_comp.sort_values('planned_courses', ascending=False)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name='مستهدف (Planned)',
            x=courses_gov_comp['المحافظة'],
            y=courses_gov_comp['planned_courses'],
            marker_color='navy',
            text=courses_gov_comp['planned_courses'],
            textposition='outside'
        ))
        fig.add_trace(go.Scatter(
            name='فعلي (Actual)',
            x=courses_gov_comp['المحافظة'],
            y=courses_gov_comp['actual_courses'],
            mode='lines+markers',
            line=dict(color='crimson', width=3),
            marker=dict(size=8)
        ))
        
        fig.update_layout(
            title='عدد الدورات / محافظة',
            xaxis_title='المحافظة',
            yaxis_title='عدد الدورات',
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.write("**عدد المتدربين / محافظة** (Number of Trainees / Governorate)")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name='مستهدف (Planned)',
            x=gov_comparison['المحافظة'],
            y=gov_comparison['عدد المستهدفين'],
            marker_color='navy',
            text=gov_comparison['عدد المستهدفين'],
            textposition='outside'
        ))
        fig.add_trace(go.Scatter(
            name='فعلي (Actual)',
            x=gov_comparison['المحافظة'],
            y=gov_comparison['actual_count'],
            mode='lines+markers',
            line=dict(color='crimson', width=3),
            marker=dict(size=8)
        ))
        
        fig.update_layout(
            title='عدد المتدربين / محافظة',
            xaxis_title='المحافظة',
            yaxis_title='عدد المتدربين',
            height=500
        )
        st.plotly_chart(fig, use_container_width=True)
    
    # Comparison table
    st.subheader("📋 Detailed Comparison Table")
    st.dataframe(
        comparison[['البرنامج التدريبي', 'عدد المستهدفين', 'actual_count', 'fulfillment_rate']].rename(columns={
            'البرنامج التدريبي': 'Program',
            'عدد المستهدفين': 'Target',
            'actual_count': 'Actual',
            'fulfillment_rate': 'Fulfillment Rate (%)'
        }),
        use_container_width=True,
        height=400
    )

# Footer
st.markdown("---")
st.markdown("**📊 Q2 Training Dashboard** | Data Source: second quarter.xlsx")
