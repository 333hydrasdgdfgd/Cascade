import os
import io
import re
import urllib.parse
import base64
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator

from django.shortcuts import render, redirect
from django.urls import reverse
from django.core.files.storage import FileSystemStorage
from django.core.files.base import ContentFile
from django.conf import settings
from django.http import HttpResponse
from django.contrib.auth import login, logout
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib.auth.decorators import login_required
from django.contrib.staticfiles import finders

ALLOWED_EXTENSIONS = ('.csv', '.xlsx', '.xls')

CHART_TYPE_LABELS = {
    'bar': 'Bar Chart',
    'horizontal_bar': 'Horizontal Bar Chart',
    'pie': 'Pie Chart',
    'line': 'Line Chart',
    'area': 'Area Chart',
    'scatter': 'Scatter Plot',
    'histogram': 'Histogram',
    'box': 'Box Plot',
    'heatmap': 'Correlation Heatmap',
}

# Time-grouping options for charts whose X axis is a date column.
# key -> (label shown in the dropdown, pandas period frequency, noun used in titles)
TIME_PERIODS = {
    'none': ('No grouping (raw rows)', None, ''),
    'D': ('Daily', 'D', 'day'),
    'W': ('Weekly', 'W', 'week'),
    'M': ('Monthly', 'M', 'month'),
    'Q': ('Quarterly', 'Q', 'quarter'),
    'Y': ('Yearly', 'Y', 'year'),
}
# key -> (label shown in the dropdown, wording used in titles)
TIME_AGGS = {
    'sum': ('Sum (total)', 'Total'),
    'mean': ('Average', 'Average'),
    'count': ('Count of records', 'Number of records'),
    'max': ('Maximum', 'Maximum'),
    'min': ('Minimum', 'Minimum'),
}
# Chart types that can show data grouped by period. (Bar charts are left out:
# draw_chart() caps them at 20 bars, which would cut off a long timeline.)
TIME_GROUPABLE_TYPES = ('line', 'area')


def landing_page(request):
    if request.user.is_authenticated:
        return redirect('upload_csv')
    # Add the 'cascade_app/' prefix here
    return render(request, 'cascade_app/Landing.html')
# --- AUTHENTICATION VIEWS ---

def login_view(request):
    if request.user.is_authenticated:
        return redirect('upload_csv')
        
    if request.method == 'POST':
        form = AuthenticationForm(request, data=request.POST)
        if form.is_valid():
            user = form.get_user()
            login(request, user)
            next_url = request.GET.get('next', 'upload_csv')
            return redirect(next_url)
    else:
        form = AuthenticationForm()
    return render(request, 'cascade_app/login.html', {'form': form})


def signup_view(request):
    if request.user.is_authenticated:
        return redirect('upload_csv')
        
    if request.method == 'POST':
        form = UserCreationForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(request, user)  # Auto-login after registration
            return redirect('upload_csv')
    else:
        form = UserCreationForm()
    return render(request, 'cascade_app/signup.html', {'form': form})


def logout_view(request):
    logout(request)
    return redirect('login')


# --- PROTECTED APP VIEWS ---

@login_required(login_url='login')
def upload_csv(request):
    if request.method == 'POST' and request.FILES.get('csv_file'):
        csv_file = request.FILES['csv_file']
        ext = os.path.splitext(csv_file.name)[1].lower()

        if ext not in ALLOWED_EXTENSIONS:
            return render(request, 'cascade_app/upload.html', {
                'error': 'Unsupported file type. Please upload a .csv, .xlsx, or .xls file.'
            })

        fs = FileSystemStorage()
        filename = fs.save(csv_file.name, csv_file)
        request.session['file_name'] = filename
        request.session.pop('insights_overrides', None)
        # A new raw upload invalidates any previously cleaned file/session data
        request.session.pop('cleaned_file_name', None)
        request.session.pop('cleaned_csv', None)
        request.session.pop('charts', None)
        return redirect('clean_data')

    return render(request, 'cascade_app/upload.html')


@login_required(login_url='login')
def visualize_data(request):
    # Visualization always runs on the cleaned dataset, never the raw upload.
    file_name = request.session.get('cleaned_file_name')
    if not file_name:
        return redirect('clean_data')

    file_path = os.path.join(settings.MEDIA_ROOT, file_name)
    if not os.path.exists(file_path):
        return redirect('clean_data')

    if 'charts' not in request.session or not request.session['charts']:
        request.session['charts'] = [{'x_col': '', 'y_col': '', 'chart_type': 'auto'}]

    if request.method == 'POST':
        action = request.POST.get('action')
        charts = request.session.get('charts', [])
        overrides = request.session.get('insights_overrides', {})
        summary_override = request.session.get('summary_override')
        recommendations_override = request.session.get('recommendations_override')
        reset_generated = False  # True when the charts change
        anchor = ''  # page section to scroll back to after the redirect

        if action == 'add_chart':
            charts.append({'x_col': '', 'y_col': '', 'chart_type': 'auto'})
            reset_generated = True
            anchor = f'chart-{len(charts) - 1}'
        elif action == 'duplicate_chart':
            idx = int(request.POST.get('chart_index', -1))
            if 0 <= idx < len(charts):
                # Append (rather than insert) so saved insights, which are
                # keyed by chart position, stay attached to the right charts.
                charts.append(charts[idx].copy())
                reset_generated = True
                anchor = f'chart-{len(charts) - 1}'
        elif action == 'remove_chart':
            idx = int(request.POST.get('chart_index', -1))
            if 0 <= idx < len(charts):
                charts.pop(idx)
                overrides.pop(str(idx), None)
                reset_generated = True
            if not charts:
                charts = [{'x_col': '', 'y_col': '', 'chart_type': 'auto'}]
            anchor = f'chart-{min(max(idx - 1, 0), len(charts) - 1)}'
        elif action == 'update_chart':
            idx = int(request.POST.get('chart_index', 0))
            anchor = f'chart-{idx}'

            if 0 <= idx < len(charts):
                old_chart = charts[idx].copy()

                new_x_col = request.POST.get('x_col', '')
                new_y_col = request.POST.get('y_col', '')
                new_chart_type = request.POST.get('chart_type', 'auto')
                new_period = request.POST.get('period', 'none')
                new_agg = request.POST.get('agg', 'sum')
                if new_period not in TIME_PERIODS:
                    new_period = 'none'
                if new_agg not in TIME_AGGS:
                    new_agg = 'sum'

                # Update chart configuration
                charts[idx]['x_col'] = new_x_col
                charts[idx]['y_col'] = new_y_col
                charts[idx]['chart_type'] = new_chart_type
                charts[idx]['period'] = new_period
                charts[idx]['agg'] = new_agg

                # Check if the chart parameters actually changed
                chart_changed = (
                    old_chart.get('x_col', '') != new_x_col
                    or old_chart.get('y_col', '') != new_y_col
                    or old_chart.get('chart_type', 'auto') != new_chart_type
                    or old_chart.get('period', 'none') != new_period
                    or old_chart.get('agg', 'sum') != new_agg
                )

                if chart_changed:
                    reset_generated = True

                    # Remove the old manually saved insight.
                    # This forces Cascade to generate a new insight
                    # based on the new chart configuration.
                    overrides.pop(str(idx), None)

                else:
                    # Only save manually edited insight if the chart
                    # configuration itself did not change.
                    edited_text = request.POST.get('insights_text', '').strip()

                    if edited_text:
                        overrides[str(idx)] = edited_text
                    else:
                        overrides.pop(str(idx), None)
        elif action == 'save_insight':
            idx = int(request.POST.get('chart_index', -1))
            anchor = f'chart-{idx}'

            if 0 <= idx < len(charts):
                edited_text = request.POST.get('insights_text', '').strip()

                if edited_text:
                    overrides[str(idx)] = edited_text
                else:
                    overrides.pop(str(idx), None)
        elif action == 'save_summary':
            edited_summary = request.POST.get('overall_summary_text', '').strip()
            summary_override = edited_summary if edited_summary else None
            anchor = 'overall-summary'
        elif action == 'reset_summary':
            summary_override = None
            anchor = 'overall-summary'
        elif action == 'save_recommendations':
            edited_recommendations_raw = request.POST.get('recommendations_text', '')
            edited_recommendations = [
                line.strip()
                for line in edited_recommendations_raw.splitlines()
                if line.strip()
            ]
            recommendations_override = edited_recommendations if edited_recommendations else None
            anchor = 'recommendations'
        elif action == 'reset_recommendations':
            recommendations_override = None
            anchor = 'recommendations'

        # A saved summary/recommendation list describes the old charts,
        # so discard it and let the auto-generated version take over.
        if reset_generated:
            summary_override = None
            recommendations_override = None

        request.session['charts'] = charts
        request.session['insights_overrides'] = overrides
        request.session['summary_override'] = summary_override
        request.session['recommendations_override'] = recommendations_override
        request.session.modified = True
        # Redirect back to the section the user was working on instead of
        # dropping them at the top of the page.
        url = reverse('visualize_data')
        if anchor:
            url += f'#{anchor}'
        return redirect(url)

    rows, cols = 0, 0
    column_names, num_cols, date_cols = [], [], []
    rendered_charts = []
    preview_html = ''

    try:
        df = read_dataset(file_path)
        rows, cols = df.shape
        column_names = df.columns.tolist()
        num_cols = df.select_dtypes(include='number').columns.tolist()
        date_cols = [c for c in column_names if _looks_like_date(df[c])]

        preview_html = df.head(10).to_html(classes='preview-table', index=False, escape=True)
        overrides = request.session.get('insights_overrides', {})
        charts_config = request.session.get('charts', [])

        for idx, config in enumerate(charts_config):
            x_col, y_col, _, _ = _resolve_columns(df, config.get('x_col', ''), config.get('y_col', ''))
            chart_type = config.get('chart_type', 'auto')
            period = config.get('period', 'none')
            agg = config.get('agg', 'sum')

            # Group a date X axis into days/weeks/months/quarters/years if requested
            plot_df, grouping_note, draw_type = prepare_time_grouping(df, x_col, y_col, config)
            if plot_df is df:
                draw_type = chart_type

            plt.clf()
            fig, ax = plt.subplots(figsize=(10, 4.5))
            active_chart_type = draw_chart(ax, fig, plot_df, x_col, y_col, draw_type, num_cols)
            if grouping_note:
                ax.set_title(f"{ax.get_title()} ({grouping_note})")
            auto_label = CHART_TYPE_LABELS.get(active_chart_type, active_chart_type) if chart_type == 'auto' else None
            plt.tight_layout()

            buf = io.BytesIO()
            fig.savefig(buf, format='png', bbox_inches='tight', dpi=100)
            buf.seek(0)
            uri = 'data:image/png;base64,' + urllib.parse.quote(base64.b64encode(buf.read()))
            plot_html = f'<img src="{uri}" alt="Chart {idx+1}" style="max-width: 100%; height: auto;">'
            plt.close('all')

            insights_text = overrides.get(str(idx))
            if insights_text is None:
                insights_text = generate_insights(plot_df, x_col, y_col, active_chart_type, num_cols)

            rendered_charts.append({
                'index': idx,
                'x_col': x_col,
                'y_col': y_col,
                'chart_type': chart_type,
                'period': period,
                'agg': agg,
                'active_chart_type': active_chart_type,
                'auto_detected_label': auto_label,
                'plot_html': plot_html,
                'insights_text': insights_text,
                'insight_is_custom': str(idx) in overrides,
            })
    except Exception as e:
        plt.close('all')
        return HttpResponse(
            f"Error generating charts: {str(e)}",
            status=500
        )
    finally:
        plt.close('all')
    author_name = (
        request.user.get_full_name()
        or request.user.username
    )

    summary_override = request.session.get('summary_override')
    recommendations_override = request.session.get('recommendations_override')

    overall_summary = summary_override or generate_overall_summary(
        df,
        rendered_charts
    )

    recommendations = recommendations_override or generate_recommendations(
        df,
        rendered_charts
    )
    context = {
        'file_name': file_name,
        'rendered_charts': rendered_charts,
        'column_names': column_names,
        'numeric_columns': num_cols,
        'date_columns': date_cols,
        'time_periods': [(k, val[0]) for k, val in TIME_PERIODS.items()],
        'time_aggs': [(k, val[0]) for k, val in TIME_AGGS.items()],
        'preview_html': preview_html,
        'author_name': author_name,
        'rows': rows,
        'cols': cols,
        'overall_summary': overall_summary,
        'recommendations': recommendations,
        'summary_is_custom': bool(summary_override),
        'recommendations_is_custom': bool(recommendations_override),
    }
    return render(request, 'cascade_app/visualize.html', context)
def generate_overall_summary(df, rendered_charts):
    """
    Creates an overall summary from the generated chart insights
    and the actual dataset.
    """

    insights = [
        chart.get('insights_text', '').strip()
        for chart in rendered_charts
        if chart.get('insights_text', '').strip()
    ]

    if not insights:
        return "No chart insights are available yet."

    summary_parts = []

    # Dataset overview
    rows, cols = df.shape
    numeric_cols = df.select_dtypes(include='number').columns.tolist()
    categorical_cols = df.select_dtypes(exclude='number').columns.tolist()

    summary_parts.append(
        f"The dataset contains {rows:,} records across {cols} columns."
    )

    if numeric_cols:
        summary_parts.append(
            f"It contains {len(numeric_cols)} numerical "
            f"and {len(categorical_cols)} categorical column(s)."
        )

    # Add chart findings
    for insight in insights:
        summary_parts.append(insight)

    return " ".join(summary_parts)


def _iqr_outlier_count(series):
    """Number of values outside the 1.5 x IQR fences (0 if not computable)."""
    series = pd.to_numeric(series, errors='coerce').dropna()
    if len(series) < 4:
        return 0
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    if iqr == 0:
        return 0
    lower = q1 - 1.5 * iqr
    upper = q3 + 1.5 * iqr
    return int(((series < lower) | (series > upper)).sum())


def _active_chart_type(df, chart):
    """The chart type that was actually drawn (resolves 'auto')."""
    chart_type = chart.get('active_chart_type') or chart.get('chart_type') or 'auto'
    if chart_type == 'auto':
        try:
            chart_type = detect_chart_type(df, chart.get('x_col'), chart.get('y_col'))
        except Exception:
            chart_type = 'bar'
    return chart_type


def _chart_ref(chart, active_type):
    """Short, readable chart label, e.g. 'Chart 1 - Bar Chart (product by region)'."""
    label = CHART_TYPE_LABELS.get(active_type, str(active_type).replace('_', ' ').title())
    number = chart.get('index', 0) + 1
    x_col = chart.get('x_col', '')
    y_col = chart.get('y_col', '')

    if active_type == 'heatmap':
        return f"Chart {number} - {label}"
    if active_type in ('histogram', 'box') and x_col == y_col:
        return f"Chart {number} - {label} ({x_col})"
    return f"Chart {number} - {label} ({y_col} by {x_col})"


def _recommend_for_chart(df, chart, active_type, num_cols):
    """
    Recommendations that depend on what this specific chart shows.
    The numbers mirror what draw_chart() plots (e.g. bars use totals).
    """
    recs = []
    x_col = chart.get('x_col')
    y_col = chart.get('y_col')

    # If the chart is grouped by period, advise on what is actually plotted.
    try:
        df, _, _ = prepare_time_grouping(df, x_col, y_col, chart)
    except Exception:
        pass

    y_numeric = y_col in df.columns and df[y_col].dtype.kind in 'bifc'

    # -----------------------------------------------------------------
    # HISTOGRAM
    # -----------------------------------------------------------------
    if active_type == 'histogram':
        series = pd.to_numeric(df[x_col], errors='coerce').dropna()
        if len(series) < 4:
            return recs

        mean, median, skew = series.mean(), series.median(), series.skew()

        if skew > 0.5 or skew < -0.5:
            side = "high" if skew > 0 else "low"
            recs.append(
                f"The distribution is skewed toward {side} values, so "
                f"describe '{x_col}' with the median ({median:,.2f}) rather than "
                f"the mean ({mean:,.2f}) and check the extreme {side} values "
                f"for errors or special cases."
            )
        else:
            recs.append(
                f"The distribution is fairly symmetric, so the mean "
                f"({mean:,.2f}) is a reliable summary of typical '{x_col}' "
                f"values. Focus on whether the spread (standard deviation "
                f"{series.std():,.2f}) is acceptable for your goals."
            )

        n_out = _iqr_outlier_count(series)
        if n_out:
            recs.append(
                f"{n_out:,} value(s) lie far from the bulk of the data. "
                f"Verify them before using this distribution for decisions."
            )

    # -----------------------------------------------------------------
    # SCATTER PLOT
    # -----------------------------------------------------------------
    elif active_type == 'scatter':
        if x_col == y_col:
            recs.append(
                f"Both axes use the same column. Choose two different "
                f"numeric columns to reveal a relationship."
            )
            return recs

        clean = df[[x_col, y_col]].apply(pd.to_numeric, errors='coerce').dropna()
        if len(clean) < 3:
            return recs

        corr = clean[x_col].corr(clean[y_col])
        if pd.isna(corr):
            return recs

        direction = "positive" if corr > 0 else "negative"
        if abs(corr) >= 0.7:
            recs.append(
                f"'{x_col}' and '{y_col}' have a strong {direction} "
                f"relationship (r = {corr:.2f}). Investigate what drives it and "
                f"consider using one to anticipate the other, while remembering "
                f"that correlation does not prove cause."
            )
        elif abs(corr) >= 0.4:
            recs.append(
                f"The {direction} relationship between '{x_col}' and "
                f"'{y_col}' is moderate (r = {corr:.2f}). Other factors also "
                f"matter, so test additional columns or segment the points by "
                f"a category."
            )
        else:
            recs.append(
                f"'{x_col}' and '{y_col}' show little linear relationship "
                f"(r = {corr:.2f}). Try other column pairs, or segment the data "
                f"by a category to look for hidden patterns."
            )

    # -----------------------------------------------------------------
    # BAR / HORIZONTAL BAR / PIE
    # -----------------------------------------------------------------
    elif active_type in ('bar', 'horizontal_bar', 'pie'):
        counts = None
        if y_numeric and x_col != y_col:
            grouped = df.groupby(x_col)[y_col].sum()
            counts = df.groupby(x_col)[y_col].count()
            measure = f"total '{y_col}'"
        else:
            grouped = df[x_col].value_counts()
            measure = "record count"

        grouped = grouped.dropna()
        n_groups = len(grouped)
        if n_groups < 2:
            return recs

        if df[x_col].dtype.kind in 'bifc' and n_groups > 20:
            shape = "slice" if active_type == 'pie' else "bar"
            recs.append(
                f"'{x_col}' is numeric with {n_groups:,} different "
                f"values, so each value becomes its own {shape} and no pattern "
                f"can be seen. Use a histogram to see how '{x_col}' is "
                f"distributed, or group the values into ranges first."
            )
            return recs

        total = grouped.sum()
        shares_valid = total > 0 and bool((grouped >= 0).all())
        top = grouped.idxmax()
        low = grouped.idxmin()
        top_share = grouped.max() / total * 100 if shares_valid else None

        if top_share is not None and top_share >= 40:
            recs.append(
                f"'{top}' makes up {top_share:.0f}% of the {measure}. "
                f"Look into what drives this concentration and whether relying "
                f"so heavily on one {x_col} value is a risk."
            )
        else:
            recs.append(
                f"'{top}' is highest and '{low}' is lowest in {measure}. "
                f"Study what separates them and decide whether '{low}' needs "
                f"attention or '{top}' practices can be replicated."
            )

        if active_type == 'pie':
            if n_groups > 7 and shares_valid:
                ordered = grouped.sort_values(ascending=False)
                other_share = ordered.iloc[7:].sum() / total * 100
                recs.append(
                    f"{n_groups} categories are squeezed into 7 slices plus "
                    f"'Other' ({other_share:.0f}%). Use a bar chart or filter the "
                    f"categories so smaller groups are not hidden."
                )
            elif shares_valid and n_groups >= 2:
                ordered = grouped.sort_values(ascending=False)
                gap = (ordered.iloc[0] - ordered.iloc[1]) / total * 100
                if gap < 5:
                    recs.append(
                        f"The top two slices are within {gap:.1f} points of "
                        f"each other. Use a bar chart if you need to rank them "
                        f"precisely."
                    )
        else:
            limit = 20 if active_type == 'bar' else 15
            if n_groups > limit:
                recs.append(
                    f"The chart shows only {limit} of {n_groups} "
                    f"'{x_col}' categories. Filter or group the rest so "
                    f"important categories are not left out."
                )

            if counts is not None and counts.min() > 0 and counts.max() >= 3 * counts.min():
                recs.append(
                    f"Bars show totals, but groups have very different "
                    f"numbers of records ({int(counts.min()):,} to "
                    f"{int(counts.max()):,}). Compare the average '{y_col}' per "
                    f"record too before deciding which '{x_col}' performs best."
                )

    # -----------------------------------------------------------------
    # LINE / AREA (trend charts)
    # -----------------------------------------------------------------
    elif active_type in ('line', 'area'):
        if x_col == y_col:
            return recs

        clean = df[[x_col, y_col]].copy()
        clean[y_col] = pd.to_numeric(clean[y_col], errors='coerce')
        clean = clean.dropna()
        if len(clean) < 3:
            return recs

        # The chart connects rows in file order.
        try:
            x_dates = pd.to_datetime(clean[x_col], errors='coerce')
            if x_dates.notna().mean() > 0.8 and not x_dates.dropna().is_monotonic_increasing:
                recs.append(
                    f"Rows are not in chronological order, so the line "
                    f"follows file order rather than time. Sort the data by "
                    f"'{x_col}' before reading the trend."
                )
        except Exception:
            pass

        values = clean[y_col].values
        first, last = values[0], values[-1]
        peak_pos = int(np.argmax(values))
        low_pos = int(np.argmin(values))
        peak_x = clean[x_col].iloc[peak_pos]
        low_x = clean[x_col].iloc[low_pos]

        if first != 0:
            change = (last - first) / abs(first) * 100
            if change >= 10:
                recs.append(
                    f"'{y_col}' rose from {first:,.2f} to {last:,.2f} "
                    f"({change:+.1f}%). Identify what is driving the growth and "
                    f"whether it can be sustained."
                )
            elif change <= -10:
                recs.append(
                    f"'{y_col}' fell from {first:,.2f} to {last:,.2f} "
                    f"({change:+.1f}%). Determine the cause of the decline and "
                    f"whether corrective action is needed."
                )
            else:
                recs.append(
                    f"'{y_col}' is fairly stable overall ({change:+.1f}% "
                    f"from start to end). Focus on the peak at '{peak_x}' "
                    f"({values[peak_pos]:,.2f}) and the low at '{low_x}' "
                    f"({values[low_pos]:,.2f}) to find what caused them."
                )
        else:
            recs.append(
                f"'{y_col}' peaks at '{peak_x}' ({values[peak_pos]:,.2f}) "
                f"and is lowest at '{low_x}' ({values[low_pos]:,.2f}). "
                f"Investigate what happened in those periods."
            )

        mean = values.mean()
        if mean != 0 and values.std() / abs(mean) > 0.5:
            recs.append(
                f"Values swing widely around the average. Check for "
                f"seasonality or one-off events around '{peak_x}' and '{low_x}' "
                f"before forecasting from this trend."
            )

    # -----------------------------------------------------------------
    # BOX PLOT
    # -----------------------------------------------------------------
    elif active_type == 'box':
        if x_col != y_col and y_numeric:
            groups = df.groupby(x_col)[y_col]
            medians = groups.median().dropna()

            if len(medians) >= 2:
                hi, lo = medians.idxmax(), medians.idxmin()
                recs.append(
                    f"'{hi}' has the highest median '{y_col}' "
                    f"({medians.max():,.2f}) and '{lo}' the lowest "
                    f"({medians.min():,.2f}). Investigate why these groups differ."
                )

            affected = 0
            total_out = 0
            for _, series in groups:
                n_out = _iqr_outlier_count(series)
                if n_out:
                    affected += 1
                    total_out += n_out
            if total_out:
                recs.append(
                    f"{total_out:,} potential outlier(s) appear in "
                    f"{affected} group(s). Verify them so they do not distort "
                    f"the comparison between groups."
                )
        else:
            n_out = _iqr_outlier_count(df[x_col])
            if n_out:
                recs.append(
                    f"{n_out:,} potential outlier(s) were found in "
                    f"'{x_col}'. Confirm whether they are errors or genuine "
                    f"extreme cases."
                )

    # -----------------------------------------------------------------
    # HEATMAP
    # -----------------------------------------------------------------
    elif active_type == 'heatmap':
        if len(num_cols) < 2:
            recs.append(
                f"A correlation heatmap needs at least two numeric "
                f"columns. Convert or add numeric columns to make it useful."
            )
            return recs

        corr = df[num_cols].corr()
        pairs = []
        for i in range(len(num_cols)):
            for j in range(i + 1, len(num_cols)):
                value = corr.iloc[i, j]
                if pd.notna(value):
                    pairs.append((abs(value), value, num_cols[i], num_cols[j]))
        if not pairs:
            return recs

        pairs.sort(reverse=True)
        strong = [p for p in pairs if p[0] >= 0.7]
        _, value, col_a, col_b = pairs[0]

        if strong:
            recs.append(
                f"{len(strong)} column pair(s) are strongly correlated "
                f"(|r| ≥ 0.70); the strongest is '{col_a}' and '{col_b}' "
                f"(r = {value:.2f}). Examine these relationships and avoid "
                f"feeding heavily overlapping columns into the same model."
            )
        else:
            recs.append(
                f"No column pair is strongly correlated (strongest: "
                f"'{col_a}' and '{col_b}', r = {value:.2f}). Explore category "
                f"groupings or non-linear patterns instead."
            )

    return recs


def generate_recommendations(df, rendered_charts):
    """
    Generates recommendations that fit the charts actually displayed.
    Each chart gets advice based on its type and the columns it plots;
    data-quality advice is limited to the columns the charts use.
    """

    if not rendered_charts:
        return [
            "Create at least one chart so recommendations can be tailored "
            "to what you are analyzing."
        ]

    recommendations = []
    num_cols = df.select_dtypes(include='number').columns.tolist()
    charted_cols = []

    # 1. Chart-specific recommendations
    for chart in rendered_charts:
        active_type = _active_chart_type(df, chart)

        cols = num_cols if active_type == 'heatmap' else [chart.get('x_col'), chart.get('y_col')]
        for col in cols:
            if col in df.columns and col not in charted_cols:
                charted_cols.append(col)

        try:
            advice = _recommend_for_chart(df, chart, active_type, num_cols)
        except Exception:
            continue  # one problematic chart should not break the page

        if advice:
            recommendations.append(
                f"{_chart_ref(chart, active_type)}: " + " ".join(advice)
            )

    # 2. Missing data in the columns the charts depend on
    if charted_cols and len(df) > 0:
        missing = df[charted_cols].isnull().sum()
        missing = missing[missing > 0]

        if not missing.empty:
            col = missing.idxmax()
            count = int(missing.max())
            pct = count / len(df) * 100
            recommendations.append(
                f"Resolve missing data in '{col}' ({count:,} value(s), "
                f"{pct:.1f}% of rows). Charts using this column leave those "
                f"records out, which can distort the totals and trends shown."
            )

    # 3. Fallback
    if not recommendations:
        recommendations.append(
            "Continue monitoring the key variables shown in the charts "
            "and compare future datasets to identify meaningful changes "
            "or emerging trends."
        )

    return recommendations


@login_required(login_url='login')
def export_pdf(request):
    file_name = request.session.get('cleaned_file_name')

    if not file_name:
        return redirect('clean_data')

    file_path = os.path.join(settings.MEDIA_ROOT, file_name)

    if not os.path.exists(file_path):
        return redirect('clean_data')

    try:
        # ---------------------------------------------------------
        # READ DATASET
        # ---------------------------------------------------------
        df = read_dataset(file_path)

        rows, cols = df.shape
        num_cols = df.select_dtypes(include='number').columns.tolist()

        # ---------------------------------------------------------
        # GET SAVED CHART CONFIGURATIONS
        # ---------------------------------------------------------
        overrides = request.session.get(
            'insights_overrides',
            {}
        )

        charts_config = request.session.get(
            'charts',
            []
        )

        rendered_charts = []

        # ---------------------------------------------------------
        # GENERATE EACH CHART
        # ---------------------------------------------------------
        for idx, config in enumerate(charts_config):

            x_col, y_col, _, _ = _resolve_columns(
                df,
                config.get('x_col', ''),
                config.get('y_col', '')
            )

            chart_type = config.get(
                'chart_type',
                'auto'
            )

            period = config.get('period', 'none')
            agg = config.get('agg', 'sum')

            # Group a date X axis into days/weeks/months/quarters/years if requested
            plot_df, grouping_note, draw_type = prepare_time_grouping(
                df, x_col, y_col, config
            )
            if plot_df is df:
                draw_type = chart_type

            # Create chart
            plt.clf()

            fig, ax = plt.subplots(
                figsize=(8.5, 4.2)
            )

            active_chart_type = draw_chart(
                ax,
                fig,
                plot_df,
                x_col,
                y_col,
                draw_type,
                num_cols
            )

            if grouping_note:
                ax.set_title(f"{ax.get_title()} ({grouping_note})")

            plt.tight_layout()

            # -----------------------------------------------------
            # CONVERT CHART TO BASE64 IMAGE
            # -----------------------------------------------------
            buf = io.BytesIO()

            fig.savefig(
                buf,
                format='png',
                bbox_inches='tight',
                dpi=120
            )

            buf.seek(0)

            uri = (
                'data:image/png;base64,'
                + urllib.parse.quote(
                    base64.b64encode(
                        buf.read()
                    )
                )
            )

            plt.close('all')

            # -----------------------------------------------------
            # GENERATE / GET INSIGHT
            # -----------------------------------------------------
            insights_text = overrides.get(
                str(idx)
            )

            if insights_text is None:
                insights_text = generate_insights(
                    plot_df,
                    x_col,
                    y_col,
                    active_chart_type,
                    num_cols
                )

            # -----------------------------------------------------
            # STORE CHART INFORMATION
            # -----------------------------------------------------
            rendered_charts.append({
                'index': idx,
                'x_col': x_col,
                'y_col': y_col,
                'chart_type': chart_type,
                'period': period,
                'agg': agg,
                'active_chart_type': active_chart_type,
                'image_uri': uri,
                'insights_text': insights_text
            })


        # ---------------------------------------------------------
        # OVERALL SUMMARY
        # ---------------------------------------------------------
        summary_override = request.session.get('summary_override')
        recommendations_override = request.session.get('recommendations_override')

        overall_summary = summary_override or generate_overall_summary(
            df,
            rendered_charts
        )

        # ---------------------------------------------------------
        # RECOMMENDATIONS
        # ---------------------------------------------------------
        recommendations = recommendations_override or generate_recommendations(
            df,
            rendered_charts
        )

        # ---------------------------------------------------------
        # PDF CONTEXT
        # ---------------------------------------------------------
        author_name = (
            request.user.get_full_name()
            or request.user.username
        )

        context = {
            'file_name': file_name,

            'charts': rendered_charts,

            'author_name': author_name,

            # Dataset information
            'rows': rows,
            'cols': cols,

            # Overall analysis
            'overall_summary': overall_summary,
            'recommendations': recommendations,
        }

        return render(
            request,
            'cascade_app/pdf_template.html',
            context
        )

    except Exception as e:

        plt.close('all')

        return HttpResponse(
            f"Error generating report view: {str(e)}",
            status=500
        )

# --- HELPER FUNCTIONS ---

def read_dataset(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        return pd.read_excel(file_path)
    return pd.read_csv(file_path)

def _looks_like_date(series, sample_size=25):
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if series.dtype.kind in 'bifc':
        return False
    sample = series.dropna().astype(str).head(sample_size)
    if sample.empty:
        return False
    try:
        parsed = pd.to_datetime(sample, errors='coerce')
        return parsed.notna().mean() > 0.8
    except Exception:
        return False

def _period_label(period, freq):
    """Readable, chronologically sortable label for a pandas Period."""
    if freq == 'Q':
        return f"{period.year}-Q{period.quarter}"
    if freq == 'M':
        return period.strftime('%Y-%m')
    if freq == 'Y':
        return str(period.year)
    # Daily -> that day; Weekly -> the date the week starts on
    return period.start_time.strftime('%Y-%m-%d')


def group_by_period(df, x_col, y_col, period='M', agg='sum'):
    """Collapse rows into one row per day / week / month / quarter / year.

    x_col is parsed as dates; y_col is aggregated with `agg`
    (sum, mean, count, max or min). Periods with no records inside the
    date range are kept, so the timeline has no silent gaps: they show as
    0 for sum/count and as a break in the line for mean/max/min.

    Returns a two-column DataFrame [x_col, y_col] in chronological order.
    """
    freq = TIME_PERIODS[period][1]

    dates = pd.to_datetime(df[x_col], errors='coerce')
    values = pd.to_numeric(df[y_col], errors='coerce')
    work = pd.DataFrame({'period': dates.dt.to_period(freq), 'value': values})
    work = work.dropna(subset=['period'])
    if work.empty:
        return pd.DataFrame({x_col: [], y_col: []})

    grouped = work.groupby('period')['value']
    result = grouped.size() if agg == 'count' else getattr(grouped, agg)()

    full_range = pd.period_range(result.index.min(), result.index.max(), freq=freq)
    result = result.reindex(full_range)
    if agg in ('sum', 'count'):
        result = result.fillna(0)

    return pd.DataFrame({
        x_col: [_period_label(p, freq) for p in result.index],
        y_col: result.to_numpy(),
    })


def prepare_time_grouping(df, x_col, y_col, config):
    """Apply the chart's 'Group by' setting when it makes sense.

    Returns (dataframe_to_plot, note_for_title, chart_type_to_draw).
    If grouping does not apply (no period chosen, X is not a date, Y is not
    numeric, or the chart type cannot show a time series), the original
    `df` object is returned untouched so callers can detect that with `is`.
    """
    period = config.get('period', 'none')
    agg = config.get('agg', 'sum')
    if period not in TIME_PERIODS or period == 'none' or agg not in TIME_AGGS:
        return df, '', None

    if x_col == y_col or x_col not in df.columns or y_col not in df.columns:
        return df, '', None
    if df[y_col].dtype.kind not in 'bifc' or not _looks_like_date(df[x_col]):
        return df, '', None

    chart_type = config.get('chart_type', 'auto')
    resolved = detect_chart_type(df, x_col, y_col) if chart_type == 'auto' else chart_type
    if resolved not in TIME_GROUPABLE_TYPES:
        return df, '', None

    grouped = group_by_period(df, x_col, y_col, period, agg)
    if grouped.empty:
        return df, '', None

    note = f"{TIME_AGGS[agg][1]} per {TIME_PERIODS[period][2]}"
    # The grouped labels (e.g. "2024-Q1") no longer look like dates, so hand
    # back the chart type that was already resolved on the original data.
    return grouped, note, resolved


def detect_chart_type(df, x_col, y_col):
    x_is_numeric = df[x_col].dtype.kind in 'bifc'
    y_is_numeric = df[y_col].dtype.kind in 'bifc'
    x_is_date = _looks_like_date(df[x_col])
    x_unique = df[x_col].nunique(dropna=True)

    if x_col == y_col and x_is_numeric:
        return 'histogram'
    if x_is_date and y_is_numeric:
        return 'line'
    if x_is_numeric and y_is_numeric:
        return 'scatter'
    if not x_is_numeric and y_is_numeric:
        return 'horizontal_bar' if x_unique > 12 else 'bar'
    return 'bar'

def draw_chart(ax, fig, df, x_col, y_col, chart_type, num_cols):
    active_chart_type = detect_chart_type(df, x_col, y_col) if chart_type == 'auto' else chart_type

    if active_chart_type == 'bar':
        if df[y_col].dtype.kind in 'bifc':
            grouped = df.groupby(x_col, as_index=False)[y_col].sum().head(20)
            ax.bar(grouped[x_col].astype(str), grouped[y_col], color='#4f46e5', edgecolor='#312e81')
            ax.set_ylabel(f"Total {y_col}")
        else:
            counts = df[x_col].value_counts().head(20)
            ax.bar(counts.index.astype(str), counts.values, color='#4f46e5', edgecolor='#312e81')
            ax.set_ylabel("Count")
        ax.set_xlabel(x_col)
        ax.set_title(f"Bar Chart: {y_col} by {x_col}")
        for label in ax.get_xticklabels():
            label.set_rotation(45)
            label.set_ha('right')

    elif active_chart_type == 'horizontal_bar':
        if df[y_col].dtype.kind in 'bifc':
            grouped = df.groupby(x_col, as_index=False)[y_col].sum().head(15)
            ax.barh(grouped[x_col].astype(str), grouped[y_col], color='#0284c7', edgecolor='#075985')
            ax.set_xlabel(f"Total {y_col}")
        else:
            counts = df[x_col].value_counts().head(15)
            ax.barh(counts.index.astype(str), counts.values, color='#0284c7', edgecolor='#075985')
            ax.set_xlabel("Count")
        ax.set_ylabel(x_col)
        ax.set_title(f"Horizontal Bar Chart: {y_col} by {x_col}")
        ax.invert_yaxis()

    elif active_chart_type == 'pie':
        if df[y_col].dtype.kind in 'bifc':
            grouped = df.groupby(x_col)[y_col].sum().sort_values(ascending=False)
        else:
            grouped = df[x_col].value_counts()

        if len(grouped) > 7:
            top_grouped = grouped.iloc[:7]
            other_sum = pd.Series({'Other': grouped.iloc[7:].sum()})
            grouped = pd.concat([top_grouped, other_sum])

        colors = ['#4f46e5', '#0284c7', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899', '#94a3b8']
        ax.pie(grouped.values, labels=grouped.index.astype(str), autopct='%1.1f%%', startangle=140, colors=colors[:len(grouped)])
        ax.set_title(f"Proportion Share: {y_col} by {x_col}")

    elif active_chart_type == 'line':
        clean_df = df.dropna(subset=[x_col, y_col])
        ax.plot(clean_df[x_col].astype(str), clean_df[y_col], marker='o', color='#4f46e5', linewidth=2)
        ax.set_xlabel(x_col)
        ax.set_ylabel(y_col)
        ax.set_title(f"Line Chart: {y_col} vs {x_col}")
        if clean_df[x_col].nunique() > 24:
            ax.xaxis.set_major_locator(MaxNLocator(nbins=12, integer=True))
        for label in ax.get_xticklabels():
            label.set_rotation(45)
            label.set_ha('right')

    elif active_chart_type == 'area':
        clean_df = df.dropna(subset=[x_col, y_col])
        x_vals = clean_df[x_col].astype(str)
        y_vals = clean_df[y_col]
        ax.fill_between(range(len(x_vals)), y_vals, color='#6366f1', alpha=0.4)
        ax.plot(range(len(x_vals)), y_vals, color='#4f46e5', linewidth=2)
        tick_step = max(1, len(x_vals) // 12)
        tick_pos = list(range(0, len(x_vals), tick_step))
        ax.set_xticks(tick_pos)
        ax.set_xticklabels(x_vals.iloc[tick_pos], rotation=45, ha='right')
        ax.set_xlabel(x_col)
        ax.set_ylabel(y_col)
        ax.set_title(f"Area Trend: {y_col} over {x_col}")

    elif active_chart_type == 'scatter':
        ax.scatter(df[x_col], df[y_col], color='#0284c7', alpha=0.7, edgecolors='#0369a1')
        ax.set_xlabel(x_col)
        ax.set_ylabel(y_col)
        ax.set_title(f"Scatter Plot: {x_col} vs {y_col}")

    elif active_chart_type == 'histogram':
        ax.hist(df[x_col].dropna(), bins=20, color='#10b981', edgecolor='#047857')
        ax.set_xlabel(x_col)
        ax.set_ylabel("Frequency")
        ax.set_title(f"Distribution of {x_col}")

    elif active_chart_type == 'box':
        if x_col != y_col and df[y_col].dtype.kind in 'bifc':
            df.boxplot(column=y_col, by=x_col, ax=ax)
            ax.figure.suptitle('')
            ax.set_title(f"Box Plot of {y_col} grouped by {x_col}")
        else:
            ax.boxplot(df[x_col].dropna())
            ax.set_title(f"Box Plot of {x_col}")
        for label in ax.get_xticklabels():
            label.set_rotation(45)
            label.set_ha('right')

    elif active_chart_type == 'heatmap':
        if len(num_cols) >= 2:
            corr = df[num_cols].corr()
            cax = ax.matshow(corr, cmap='coolwarm', vmin=-1, vmax=1)
            fig.colorbar(cax, ax=ax, fraction=0.046, pad=0.04)
            ax.set_xticks(range(len(num_cols)))
            ax.set_yticks(range(len(num_cols)))
            ax.set_xticklabels(num_cols, rotation=45, ha='left')
            ax.set_yticklabels(num_cols)
            for i in range(len(num_cols)):
                for j in range(len(num_cols)):
                    val = corr.iloc[i, j]
                    ax.text(j, i, f"{val:.2f}", ha='center', va='center', color='black' if -0.5 < val < 0.5 else 'white')
            ax.set_title("Numerical Correlation Matrix", pad=20)
        else:
            ax.text(0.5, 0.5, "Heatmap requires at least 2 numerical columns.", ha='center', va='center')

    return active_chart_type

def generate_insights(df, x_col, y_col, chart_type, num_cols):
    """
    Generate insights that directly correspond to the data,
    parameters, aggregation, and chart type being displayed.
    """

    try:
        # ---------------------------------------------------------
        # HISTOGRAM
        # ---------------------------------------------------------
        if chart_type == 'histogram':
            series = pd.to_numeric(df[x_col], errors='coerce').dropna()

            if series.empty:
                return f"No valid numeric data was found for {x_col}."

            mean = series.mean()
            median = series.median()
            std = series.std()
            min_val = series.min()
            max_val = series.max()

            skew = series.skew()

            if skew > 0.5:
                shape = "right-skewed"
            elif skew < -0.5:
                shape = "left-skewed"
            else:
                shape = "approximately symmetric"

            return (
                f"{x_col} has a mean of {mean:.2f} and a median of {median:.2f}. "
                f"Values range from {min_val:.2f} to {max_val:.2f}, "
                f"with a standard deviation of {std:.2f}. "
                f"The distribution appears {shape}."
            )

        # ---------------------------------------------------------
        # SCATTER PLOT
        # ---------------------------------------------------------
        if chart_type == 'scatter':
            clean = df[[x_col, y_col]].dropna()

            if len(clean) < 2:
                return (
                    f"Not enough valid observations to determine the "
                    f"relationship between {x_col} and {y_col}."
                )

            corr = clean[x_col].corr(clean[y_col])

            if pd.isna(corr):
                return (
                    f"A correlation could not be calculated between "
                    f"{x_col} and {y_col}."
                )

            abs_corr = abs(corr)

            if abs_corr >= 0.7:
                strength = "strong"
            elif abs_corr >= 0.4:
                strength = "moderate"
            else:
                strength = "weak"

            if corr > 0:
                direction = "positive"
            elif corr < 0:
                direction = "negative"
            else:
                direction = "no clear"

            if corr > 0:
                explanation = (
                    f"Higher values of {x_col} generally correspond "
                    f"with higher values of {y_col}."
                )
            elif corr < 0:
                explanation = (
                    f"Higher values of {x_col} generally correspond "
                    f"with lower values of {y_col}."
                )
            else:
                explanation = (
                    f"There is no clear linear relationship between "
                    f"{x_col} and {y_col}."
                )

            return (
                f"{x_col} and {y_col} show a {strength} {direction} "
                f"relationship (correlation: {corr:.2f}). "
                f"{explanation}"
            )

        # ---------------------------------------------------------
        # LINE CHART
        # ---------------------------------------------------------
        if chart_type == 'line':
            clean = df[[x_col, y_col]].dropna()

            if len(clean) < 2:
                return (
                    f"Not enough valid observations to identify a trend "
                    f"between {x_col} and {y_col}."
                )

            values = pd.to_numeric(clean[y_col], errors='coerce')
            valid = values.notna()

            clean = clean.loc[valid].copy()
            values = values.loc[valid]

            if len(clean) < 2:
                return f"No valid numeric values were found for {y_col}."

            first = values.iloc[0]
            last = values.iloc[-1]

            change = last - first

            if first != 0:
                pct_change = (change / abs(first)) * 100
            else:
                pct_change = None

            if change > 0:
                trend = "increased"
            elif change < 0:
                trend = "decreased"
            else:
                trend = "remained unchanged"

            max_index = values.idxmax()
            min_index = values.idxmin()

            peak_x = clean.loc[max_index, x_col]
            lowest_x = clean.loc[min_index, x_col]

            if pct_change is not None:
                change_text = f"{abs(pct_change):.1f}%"
            else:
                change_text = "from a zero starting value"

            return (
                f"{y_col} {trend} from {first:.2f} to {last:.2f}, "
                f"a change of {change_text}. "
                f"The highest value was {values.max():.2f} at "
                f"{x_col} = {peak_x}, while the lowest value was "
                f"{values.min():.2f} at {x_col} = {lowest_x}."
            )

        # ---------------------------------------------------------
        # AREA CHART
        # ---------------------------------------------------------
        if chart_type == 'area':
            clean = df[[x_col, y_col]].dropna()

            if clean.empty:
                return f"No valid data was found for {x_col} and {y_col}."

            values = pd.to_numeric(clean[y_col], errors='coerce').dropna()

            if values.empty:
                return f"No valid numeric values were found for {y_col}."

            total = values.sum()
            average = values.mean()
            maximum = values.max()
            minimum = values.min()

            max_position = values.idxmax()
            min_position = values.idxmin()

            max_x = clean.loc[max_position, x_col]
            min_x = clean.loc[min_position, x_col]

            return (
                f"The area chart shows {y_col} across {len(values)} "
                f"observations of {x_col}. The total value is {total:.2f}, "
                f"with an average of {average:.2f}. "
                f"The highest value was {maximum:.2f} at "
                f"{x_col} = {max_x}, while the lowest was "
                f"{minimum:.2f} at {x_col} = {min_x}."
            )

        # ---------------------------------------------------------
        # BAR / HORIZONTAL BAR
        # ---------------------------------------------------------
        if chart_type in ('bar', 'horizontal_bar'):

            # Numeric Y = grouped SUM
            if df[y_col].dtype.kind in 'bifc':

                grouped = (
                    df.groupby(x_col)[y_col]
                    .sum()
                    .sort_values(ascending=False)
                )

                if grouped.empty:
                    return (
                        f"No grouped values could be calculated for "
                        f"{y_col} by {x_col}."
                    )

                total = grouped.sum()

                highest_category = grouped.index[0]
                highest_value = grouped.iloc[0]

                lowest_category = grouped.index[-1]
                lowest_value = grouped.iloc[-1]

                if total != 0:
                    highest_share = (highest_value / total) * 100
                else:
                    highest_share = 0

                return (
                    f"'{highest_category}' has the highest total {y_col} "
                    f"at {highest_value:.2f}, representing "
                    f"{highest_share:.1f}% of the overall total. "
                    f"'{lowest_category}' has the lowest total at "
                    f"{lowest_value:.2f}."
                )

            # Non-numeric Y = frequency COUNT
            counts = df[x_col].value_counts()

            if counts.empty:
                return f"No categorical values were found in {x_col}."

            total = counts.sum()

            highest_category = counts.index[0]
            highest_count = counts.iloc[0]

            lowest_category = counts.index[-1]
            lowest_count = counts.iloc[-1]

            share = (highest_count / total) * 100 if total else 0

            return (
                f"'{highest_category}' is the most frequent value in "
                f"{x_col}, appearing {highest_count} times "
                f"({share:.1f}% of {total} records). "
                f"'{lowest_category}' appears the fewest times, "
                f"with {lowest_count} occurrence(s)."
            )

        # ---------------------------------------------------------
        # PIE CHART
        # ---------------------------------------------------------
        if chart_type == 'pie':

            # Match the pie chart's actual grouping logic
            if df[y_col].dtype.kind in 'bifc':
                grouped = (
                    df.groupby(x_col)[y_col]
                    .sum()
                    .sort_values(ascending=False)
                )
            else:
                grouped = df[x_col].value_counts()

            if grouped.empty:
                return f"No data was available to create proportions."

            total = grouped.sum()

            highest_category = grouped.index[0]
            highest_value = grouped.iloc[0]

            share = (highest_value / total) * 100 if total else 0

            # Match the chart behavior where categories after
            # the first 7 are combined into "Other"
            if len(grouped) > 7:
                displayed_categories = list(grouped.index[:7])
                displayed_categories.append("Other")

                other_value = grouped.iloc[7:].sum()

                if highest_category == "Other":
                    highest_category = "Other"
                    highest_value = other_value
                    share = (
                        (other_value / total) * 100
                        if total else 0
                    )

            return (
                f"'{highest_category}' represents the largest share "
                f"at {share:.1f}% of the total {y_col} distribution "
                f"across {len(grouped)} categories."
            )

        # ---------------------------------------------------------
        # BOX PLOT
        # ---------------------------------------------------------
        if chart_type == 'box':

            if x_col != y_col and df[y_col].dtype.kind in 'bifc':
                series = df[y_col].dropna()
                measured_column = y_col
            else:
                series = df[x_col].dropna()
                measured_column = x_col

            if series.empty:
                return f"No valid data was found for {measured_column}."

            q1 = series.quantile(0.25)
            median = series.quantile(0.50)
            q3 = series.quantile(0.75)

            iqr = q3 - q1

            lower_bound = q1 - (1.5 * iqr)
            upper_bound = q3 + (1.5 * iqr)

            outliers = series[
                (series < lower_bound) |
                (series > upper_bound)
            ]

            return (
                f"{measured_column} has a median of {median:.2f}, "
                f"with the middle 50% of values ranging from "
                f"{q1:.2f} to {q3:.2f} (IQR: {iqr:.2f}). "
                f"{len(outliers)} potential outlier(s) were detected "
                f"using the 1.5×IQR rule."
            )

        # ---------------------------------------------------------
        # CORRELATION HEATMAP
        # ---------------------------------------------------------
        if chart_type == 'heatmap':

            if len(num_cols) < 2:
                return (
                    "Not enough numerical columns are available "
                    "to calculate correlations."
                )

            corr = df[num_cols].corr()

            # Remove self-correlations from consideration
            abs_corr = corr.abs().copy()
            np.fill_diagonal(abs_corr.values, 0)

            i, j = np.unravel_index(
                np.argmax(abs_corr.values),
                abs_corr.shape
            )

            col_a = num_cols[i]
            col_b = num_cols[j]

            correlation = corr.loc[col_a, col_b]

            if pd.isna(correlation):
                return (
                    "A meaningful correlation could not be determined "
                    "from the numerical columns."
                )

            abs_value = abs(correlation)

            if abs_value >= 0.7:
                strength = "strong"
            elif abs_value >= 0.4:
                strength = "moderate"
            else:
                strength = "weak"

            if correlation > 0:
                direction = "positive"
            elif correlation < 0:
                direction = "negative"
            else:
                direction = "no"

            return (
                f"The strongest relationship in the correlation matrix "
                f"is between {col_a} and {col_b}, with a {strength} "
                f"{direction} correlation of {correlation:.2f}."
            )

        # ---------------------------------------------------------
        # FALLBACK
        # ---------------------------------------------------------
        return (
            f"No specific insight rule is available for the "
            f"{chart_type} chart."
        )

    except Exception as e:
        return (
            f"Unable to generate an insight for this chart: {str(e)}"
        )
    
def _resolve_columns(df, x_col_param, y_col_param):
    column_names = df.columns.tolist()
    num_cols = df.select_dtypes(include='number').columns.tolist()
    if not column_names:
        raise ValueError("The uploaded file contains no readable columns.")
    default_x = column_names[0]
    default_y = num_cols[0] if num_cols else (column_names[1] if len(column_names) > 1 else default_x)
    x_col = x_col_param if x_col_param in column_names else default_x
    y_col = y_col_param if y_col_param in column_names else default_y
    return x_col, y_col, column_names, num_cols

def link_callback(uri, rel):
    if os.path.isabs(uri) and os.path.exists(uri):
        return uri
    if uri.startswith(settings.MEDIA_URL):
        path = os.path.join(settings.MEDIA_ROOT, uri.replace(settings.MEDIA_URL, ""))
    elif uri.startswith(settings.STATIC_URL):
        path = os.path.join(settings.STATIC_ROOT, uri.replace(settings.STATIC_URL, ""))
    else:
        path = finders.find(uri) or uri
    return path if os.path.isfile(path) else uri
def _normalize_whitespace(df):
    """Trim leading/trailing whitespace and collapse repeated internal
    spaces in every text column. Generic — no column names assumed."""
    changed = 0
    for column in df.select_dtypes(include=["object", "str"]).columns:
        before = df[column].copy()
        df[column] = df[column].apply(
            lambda x: re.sub(r"\s+", " ", x.strip()) if isinstance(x, str) else x
        )
        changed += int((before != df[column]).sum())
    return df, changed


def _normalize_categorical_casing(df, max_unique_ratio=0.5):
    """For any text column that behaves like a category (few distinct
    values relative to row count), merge values that are identical except
    for case/whitespace into one canonical spelling — the spelling that
    occurs most often in the data. This is data-driven: it never assumes
    what the column means or what its 'correct' values should be, so it
    works the same way on any dataset."""
    total_changed = 0
    columns_affected = 0

    for column in df.select_dtypes(include=["object", "str"]).columns:
        series = df[column].dropna().astype(str)
        if series.empty:
            continue

        n_unique = series.nunique()
        if n_unique <= 1 or n_unique / len(series) > max_unique_ratio:
            # Looks like free text (names, emails, ids) rather than a
            # category, so leave it alone.
            continue

        groups = {}
        for val in series.unique():
            key = val.lower()
            groups.setdefault(key, []).append(val)

        remap = {}
        for variants in groups.values():
            if len(variants) > 1:
                counts = series[series.isin(variants)].value_counts()
                canonical = counts.idxmax()
                for v in variants:
                    if v != canonical:
                        remap[v] = canonical

        if remap:
            before = df[column].copy()
            df[column] = df[column].apply(
                lambda x: remap.get(x, x) if isinstance(x, str) else x
            )
            n_changed = int((before != df[column]).sum())
            if n_changed > 0:
                total_changed += n_changed
                columns_affected += 1

    return df, total_changed, columns_affected


def _normalize_dates(df):
    """Detect any column whose values look like dates (reusing the same
    heuristic the chart auto-detector uses) and rewrite it to a single
    ISO 8601 format (YYYY-MM-DD), regardless of what mixed formats the
    original values were in."""
    columns_affected = 0

    for column in df.columns:
        series = df[column]
        if not (pd.api.types.is_string_dtype(series) or pd.api.types.is_datetime64_any_dtype(series)):
            continue
        if not _looks_like_date(series):
            continue

        parsed = pd.to_datetime(series, errors="coerce", format="mixed")
        non_null = series.notna().sum()
        if non_null == 0:
            continue

        if parsed.notna().sum() / non_null >= 0.8:
            df[column] = parsed.dt.strftime("%Y-%m-%d")
            columns_affected += 1

    return df, columns_affected


def _tighten_numeric_types(df):
    """After missing-value handling, a numeric column with even one blank
    gets upcast to float (e.g. 21 -> 21.0). If every non-null value in a
    numeric column is a whole number, cast it back to a nullable integer
    type so it displays as 21, not 21.0 — again, with no assumption about
    what the column represents."""
    columns_affected = 0

    for column in df.select_dtypes(include="number").columns:
        series = df[column].dropna()
        if series.empty:
            continue
        if (series % 1 == 0).all():
            try:
                df[column] = df[column].astype("Int64")
                columns_affected += 1
            except (TypeError, ValueError):
                pass

    return df, columns_affected


def _flag_outliers(df):
    """Report numeric values that fall outside the typical IQR range for
    their column. These are flagged for manual review only — never
    changed or removed automatically, since 'outlier' doesn't always mean
    'wrong' and this cleaner has no domain knowledge of the data."""
    flagged = []

    for column in df.select_dtypes(include="number").columns:
        series = df[column].dropna()
        if len(series) < 4:
            continue
        q1, q3 = series.quantile(0.25), series.quantile(0.75)
        iqr = q3 - q1
        if iqr == 0:
            continue
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        outliers = series[(series < lower) | (series > upper)]
        if len(outliers) > 0:
            flagged.append((column, len(outliers)))

    return flagged


def _flag_near_duplicate_rows(df, match_ratio=0.8, max_rows=1000):
    """Report pairs of rows that are identical across most (but not all)
    columns — a common sign of the same real-world record entered twice
    with a typo or a different date/ID. Flagged only, never removed,
    since collapsing them automatically risks deleting legitimate rows.
    Skipped on very large files to keep this a fast, synchronous step."""
    if len(df) > max_rows or len(df.columns) == 0:
        return None  # not checked

    values = df.astype(str).values
    n_cols = len(df.columns)
    pair_count = 0

    for i in range(len(values)):
        for j in range(i + 1, len(values)):
            matches = sum(1 for a, b in zip(values[i], values[j]) if a == b)
            ratio = matches / n_cols
            if match_ratio <= ratio < 1.0:
                pair_count += 1

    return pair_count


@login_required(login_url='login')
def clean_data(request):
    context = {}

    # This view no longer accepts its own file upload. It always operates on
    # the file the user already uploaded via upload_csv (session['file_name']).
    file_name = request.session.get("file_name")
    if not file_name:
        return redirect("upload_csv")

    file_path = os.path.join(settings.MEDIA_ROOT, file_name)
    if not os.path.exists(file_path):
        return redirect("upload_csv")

    context["uploaded_file_name"] = file_name

    if request.method == "POST":
        try:
            df = read_dataset(file_path)

            original_rows = len(df)
            original_columns = len(df.columns)

            cleaning_report = []

            # 1. Remove completely empty rows
            empty_rows = df.isna().all(axis=1).sum()

            if empty_rows > 0:
                df = df.dropna(how="all")
                cleaning_report.append(
                    f"Removed {empty_rows} completely empty row(s)."
                )

            # 2. Remove duplicate rows
            duplicate_count = df.duplicated().sum()

            if duplicate_count > 0:
                df = df.drop_duplicates()
                cleaning_report.append(
                    f"Removed {duplicate_count} duplicate row(s)."
                )

            # 3. Normalize whitespace (trim + collapse repeated spaces)
            df, spaces_removed = _normalize_whitespace(df)

            if spaces_removed > 0:
                cleaning_report.append(
                    f"Normalized spacing in {spaces_removed} value(s) (trimmed and collapsed extra spaces)."
                )

            # 4. Standardize missing values
            missing_values = [
                "",
                " ",
                "NA",
                "N/A",
                "na",
                "n/a",
                "null",
                "NULL",
                "-"
            ]

            missing_count = 0

            for column in df.columns:
                for value in missing_values:
                    missing_count += (df[column] == value).sum()

                df[column] = df[column].replace(
                    missing_values,
                    pd.NA
                )

            if missing_count > 0:
                cleaning_report.append(
                    f"Standardized {missing_count} missing/invalid value(s)."
                )

            # 5. Convert numeric-looking columns
            numeric_columns = 0

            for column in df.columns:

                if pd.api.types.is_string_dtype(df[column]):

                    converted = pd.to_numeric(
                        df[column]
                        .astype(str)
                        .str.replace(",", "", regex=False),
                        errors="coerce"
                    )

                    non_empty = df[column].notna().sum()

                    if non_empty > 0:
                        numeric_count = converted.notna().sum()

                        if numeric_count / non_empty >= 0.8:
                            df[column] = converted
                            numeric_columns += 1

            if numeric_columns > 0:
                cleaning_report.append(
                    f"Converted {numeric_columns} column(s) to numeric format."
                )

            # 6. Standardize column names
            old_columns = list(df.columns)

            df.columns = (
                df.columns
                .str.strip()
                .str.lower()
                .str.replace(" ", "_", regex=False)
            )

            changed_columns = sum(
                old != new
                for old, new in zip(old_columns, df.columns)
            )

            if changed_columns > 0:
                cleaning_report.append(
                    f"Standardized {changed_columns} column name(s)."
                )

            # 7. Normalize categorical value casing (data-driven, no hardcoded columns)
            df, casing_changed, casing_columns = _normalize_categorical_casing(df)

            if casing_columns > 0:
                cleaning_report.append(
                    f"Standardized inconsistent capitalization in {casing_changed} "
                    f"value(s) across {casing_columns} column(s)."
                )

            # 8. Standardize any date-like column to a single format (YYYY-MM-DD)
            df, date_columns = _normalize_dates(df)

            if date_columns > 0:
                cleaning_report.append(
                    f"Standardized date formats to YYYY-MM-DD in {date_columns} column(s)."
                )

            # 9. Restore whole-number formatting where a column is entirely integers
            df, tightened_columns = _tighten_numeric_types(df)

            if tightened_columns > 0:
                cleaning_report.append(
                    f"Restored whole-number formatting in {tightened_columns} numeric column(s)."
                )

            # 10. Check remaining missing values
            remaining_missing = int(
                df.isna().sum().sum()
            )

            if remaining_missing > 0:
                cleaning_report.append(
                    f"{remaining_missing} missing value(s) remain and were not automatically filled."
                )

            # 11. Flag possible outliers for manual review (never auto-changed)
            outlier_flags = _flag_outliers(df)

            if outlier_flags:
                flagged_desc = ", ".join(f"'{col}' ({count})" for col, count in outlier_flags)
                cleaning_report.append(
                    f"Flagged possible outliers for manual review in: {flagged_desc}. "
                    f"Values were not changed or removed."
                )

            # 12. Flag possible near-duplicate rows for manual review (never auto-removed)
            near_dupe_count = _flag_near_duplicate_rows(df)

            if near_dupe_count:
                cleaning_report.append(
                    f"Flagged {near_dupe_count} pair(s) of near-duplicate rows "
                    f"(similar across most columns) for manual review. Rows were not merged or removed."
                )

            # Nothing needed cleaning
            if not cleaning_report:
                cleaning_report.append(
                    "No cleaning was necessary. The dataset already appears clean."
                )

            # Save the cleaned data both as text (for the CSV download link)
            # and as a real file on disk (so visualize_data / export_pdf can
            # read it the same way they read the original upload).
            cleaned_csv_text = df.to_csv(index=False)
            request.session["cleaned_csv"] = cleaned_csv_text

            cleaned_filename = f"cleaned_{os.path.splitext(file_name)[0]}.csv"
            fs = FileSystemStorage()
            if fs.exists(cleaned_filename):
                fs.delete(cleaned_filename)
            fs.save(cleaned_filename, ContentFile(cleaned_csv_text.encode("utf-8")))

            request.session["cleaned_file_name"] = cleaned_filename
            # Chart/insight state from a previous cleaned file is now stale.
            request.session.pop("charts", None)
            request.session.pop("insights_overrides", None)
            request.session.modified = True

            context = {
                "uploaded_file_name": file_name,
                "cleaning_report": cleaning_report,
                "original_rows": original_rows,
                "cleaned_rows": len(df),
                "original_columns": original_columns,
                "cleaned_columns": len(df.columns),
                "remaining_missing": remaining_missing,
                "cleaned": True,
            }

        except Exception as e:
            context["error"] = f"Unable to clean file: {str(e)}"

    return render(request, "cascade_app/clean_data.html", context)

@login_required(login_url='login')
def download_cleaned_data(request):

    cleaned_csv = request.session.get("cleaned_csv")

    if not cleaned_csv:
        return HttpResponse(
            "No cleaned file available.",
            status=404
        )

    response = HttpResponse(
        cleaned_csv,
        content_type="text/csv"
    )

    response["Content-Disposition"] = (
        'attachment; filename="cascade_cleaned_data.csv"'
    )

    return response