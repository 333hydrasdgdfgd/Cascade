import os
import io
import json
import re
import urllib.parse
import base64
import zipfile
import hashlib
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from collections import namedtuple
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

ALLOWED_EXTENSIONS = ('.csv', '.xlsx', '.xls', '.json', '.jsonl', '.ndjson')

# Largest file the importer accepts. Cleaning, charting and the PDF all run
# inside a normal web request, and an Excel file expands to many times its
# size once it is loaded, so 5 MB (roughly 50,000+ rows of ordinary tabular
# data) keeps every step fast. Change this one number to raise or lower it;
# the upload page and the error messages follow automatically.
MAX_UPLOAD_MB = 5
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

# The cleaning step lets the user review rows with missing values one by one.
# Beyond this many rows that stops being practical, so only the
# "remove those rows" option is offered.
MANUAL_REVIEW_MAX_ROWS = 500

# Saved-progress (.cascade) files hold the cleaned dataset plus the settings,
# so they are allowed to be larger than a raw upload: three times the upload
# cap, and the data inside is limited when it is unpacked. See "SAVE / RESTORE
# PROGRESS" near the end of this file.
PROGRESS_FORMAT = 'cascade-progress'
PROGRESS_VERSION = 1
PROGRESS_EXTENSION = '.cascade'
MAX_PROGRESS_MB = MAX_UPLOAD_MB * 3
MAX_PROGRESS_BYTES = MAX_PROGRESS_MB * 1024 * 1024
MAX_PROGRESS_DATA_BYTES = 50 * 1024 * 1024
MAX_PROGRESS_MANIFEST_BYTES = 2 * 1024 * 1024
MAX_PROGRESS_COLUMNS = 500
MAX_SAVED_CHARTS = 50
MAX_SAVED_PIVOTS = 20
MAX_NOTE_CHARS = 5000
MAX_SUMMARY_CHARS = 20000
MAX_RECOMMENDATIONS = 100
JSON_EXTENSIONS = ('.json', '.jsonl', '.ndjson')

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
# Chart types that can show data grouped by period. A grouped bar chart keeps
# every period in time order (see draw_chart's keep_order / thin_labels).
TIME_GROUPABLE_TYPES = ('line', 'area', 'bar')

# --- Calculated Y axis (the "Advanced" button) ---
# operator -> symbol shown to the user
CALC_OPS = {'/': '÷', '*': '×', '-': '−', '+': '+'}
# 'total' = add the rows up per group, 'mean' = average each row's result
CALC_MODES = ('total', 'mean')
# Chart types that can show one calculated value per group
CALC_CHART_TYPES = ('bar', 'horizontal_bar', 'pie', 'line', 'area')
# A calculated X value is worked out row by row, so it suits these charts
X_CALC_CHART_TYPES = ('scatter', 'histogram', 'box')

# What is actually plotted: the data frame, the X/Y column names to draw,
# a note for the title, the chart type, and the X/Y calculation specs (or None)
PlotData = namedtuple('PlotData', 'df x y note draw_type x_spec y_spec ranked', defaults=('',))


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

def _format_size(num_bytes):
    if num_bytes < 1024 * 1024:
        return f"{max(1, round(num_bytes / 1024))} KB"
    return f"{num_bytes / (1024 * 1024):.2f} MB"


def _render_upload(request, error=None, status=200, import_error=None):
    return render(request, 'cascade_app/upload.html', {
        'error': error,
        'import_error': import_error,
        'max_upload_mb': MAX_UPLOAD_MB,
        'max_upload_bytes': MAX_UPLOAD_BYTES,
        'max_progress_mb': MAX_PROGRESS_MB,
        'max_progress_bytes': MAX_PROGRESS_BYTES,
    }, status=status)


@login_required(login_url='login')
def upload_csv(request):
    if request.method == 'POST':
        # A saved-progress file (from "Save Progress" on the Visualize page)
        progress_file = request.FILES.get('progress_file')
        if progress_file:
            return _import_progress(request, progress_file)

        csv_file = request.FILES.get('csv_file')

        if not csv_file:
            return _render_upload(request, 'Please choose a file to upload.')

        if csv_file.size == 0:
            return _render_upload(request, 'That file is empty. Please choose a file that contains data.')

        # Reject oversized files before anything is written to disk.
        if csv_file.size > MAX_UPLOAD_BYTES:
            return _render_upload(
                request,
                f'"{csv_file.name}" is {_format_size(csv_file.size)}, which is over the '
                f'{MAX_UPLOAD_MB} MB limit. Please upload a smaller file '
                f'(for example, remove unused rows or columns, or split it into parts).',
                status=413,
            )

        ext = os.path.splitext(csv_file.name)[1].lower()

        if ext not in ALLOWED_EXTENSIONS:
            return _render_upload(
                request,
                'Unsupported file type. Please upload a .csv, .xlsx, .xls, or .json file.'
            )

        fs = FileSystemStorage()
        filename = fs.save(csv_file.name, csv_file)

        # JSON can be shaped in many ways, so check now that it converts to a
        # table and fail here with a clear message instead of in the cleaning step.
        if ext in JSON_EXTENSIONS:
            try:
                read_dataset(fs.path(filename))
            except ValueError as e:
                fs.delete(filename)
                return _render_upload(request, str(e))
            except Exception:
                fs.delete(filename)
                return _render_upload(request, 'Could not read that file. Make sure it is valid JSON.')

        request.session['file_name'] = filename
        request.session.pop('insights_overrides', None)
        # A new raw upload invalidates any previously cleaned file/session data
        request.session.pop('cleaned_file_name', None)
        request.session.pop('cleaned_csv', None)
        request.session.pop('currency_symbols', None)
        request.session.pop('cleaning_actions', None)
        request.session.pop('cleaning_result', None)
        request.session.pop('charts', None)
        for key in ('filters', 'kpis', 'pivot', 'pivot_insight_override', 'pivot_recommendation_override',
                    'pivots', 'layout', 'summary_override', 'recommendations_override'):
            request.session.pop(key, None)
        return redirect('clean_data')

    return _render_upload(request)


@login_required(login_url='login')
def visualize_data(request):
    # Visualization always runs on the cleaned dataset, never the raw upload.
    file_name = request.session.get('cleaned_file_name')
    if not file_name:
        return redirect('clean_data')

    file_path = os.path.join(settings.MEDIA_ROOT, file_name)
    if not os.path.exists(file_path):
        return redirect('clean_data')

    if request.method == 'GET' and request.GET.get('download') == 'pivot':
        return _download_pivot(request, file_path)

    if request.method == 'GET' and request.GET.get('download') == 'progress':
        return _download_progress(request, file_path)

    if 'charts' not in request.session or not request.session['charts']:
        request.session['charts'] = [{'x_col': '', 'y_col': '', 'chart_type': 'auto'}]

    if request.method == 'POST':
        action = request.POST.get('action')
        charts = request.session.get('charts', [])
        overrides = request.session.get('insights_overrides', {})
        summary_override = request.session.get('summary_override')
        recommendations_override = request.session.get('recommendations_override')
        pivots = _session_pivots(request)
        layout = _normalize_layout(request.session.get('layout'), len(charts), len(pivots))
        reset_generated = False  # True when the charts change
        anchor = ''  # page section to scroll back to after the redirect

        if action == 'add_chart':
            charts.append({'x_col': '', 'y_col': '', 'chart_type': 'auto'})
            layout.append(f'c{len(charts) - 1}')
            reset_generated = True
            anchor = f'chart-{len(charts) - 1}'
        elif action == 'duplicate_chart':
            idx = _to_int(request.POST.get('chart_index'))
            if 0 <= idx < len(charts):
                # Saved at the end of the list (insights are keyed by chart position),
                # but shown right after the chart it was copied from.
                charts.append(charts[idx].copy())
                layout.insert(layout.index(f'c{idx}') + 1, f'c{len(charts) - 1}')
                reset_generated = True
                anchor = f'chart-{len(charts) - 1}'
        elif action == 'remove_chart':
            idx = _to_int(request.POST.get('chart_index'))
            if 0 <= idx < len(charts):
                charts.pop(idx)
                # charts after the removed one move up a place, and so do their saved insights
                overrides = {
                    str(int(k) - 1 if int(k) > idx else int(k)): v
                    for k, v in overrides.items() if str(k).isdigit() and int(k) != idx
                }
                layout = [
                    f'c{int(tok[1:]) - 1}' if tok[0] == 'c' and int(tok[1:]) > idx else tok
                    for tok in layout if tok != f'c{idx}'
                ]
                reset_generated = True
            if not charts:
                charts = [{'x_col': '', 'y_col': '', 'chart_type': 'auto'}]
                layout.append('c0')
            anchor = f'chart-{min(max(idx - 1, 0), len(charts) - 1)}'
        elif action == 'add_pivot':
            try:
                df_check = read_dataset(file_path)
                new_cfg = _next_pivot_config(
                    df_check, [normalize_pivot_config(df_check, p.get('cfg')) for p in pivots])
            except Exception:
                new_cfg = None
            pivots.append({'cfg': new_cfg, 'insight': None, 'recommendation': None})
            layout.append(f'p{len(pivots) - 1}')
            anchor = f'pivot-{len(pivots) - 1}'
        elif action == 'duplicate_pivot':
            pidx = _to_int(request.POST.get('pivot_index'))
            if 0 <= pidx < len(pivots):
                pivots.append({'cfg': dict(pivots[pidx]['cfg']) if pivots[pidx]['cfg'] else None,
                               'insight': None, 'recommendation': None})
                layout.insert(layout.index(f'p{pidx}') + 1, f'p{len(pivots) - 1}')
                anchor = f'pivot-{len(pivots) - 1}'
        elif action == 'remove_pivot':
            pidx = _to_int(request.POST.get('pivot_index'))
            if 0 <= pidx < len(pivots):
                pivots.pop(pidx)
                layout = [
                    f'p{int(tok[1:]) - 1}' if tok[0] == 'p' and int(tok[1:]) > pidx else tok
                    for tok in layout if tok != f'p{pidx}'
                ]
            anchor = 'filters'
        elif action in ('move_block', 'set_layout'):
            old_charts = _chart_order(layout)
            if action == 'move_block':
                token = request.POST.get('block', '')
                if token in layout and request.POST.get('direction') in ('up', 'down'):
                    here = layout.index(token)
                    there = here - 1 if request.POST.get('direction') == 'up' else here + 1
                    if 0 <= there < len(layout):
                        layout[here], layout[there] = layout[there], layout[here]
                anchor = _block_anchor(token)
            else:
                layout = _normalize_layout(request.POST.get('order', '').split(','), len(charts), len(pivots))
                anchor = _block_anchor(request.POST.get('moved', ''))
            # saved recommendations say "Chart N": keep them with the chart they were written for
            recommendations_override = _renumber_chart_refs(recommendations_override, old_charts, _chart_order(layout))
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
                new_sort, new_topn, new_topn_n = _read_rank_post(request.POST)

                # Advanced: calculated Y / X values (None when switched off)
                new_calc = _read_calc_post(request.POST, 'calc')
                new_x_calc = _read_calc_post(request.POST, 'xcalc')
                if new_calc and new_x_calc:
                    new_x_calc = None  # one calculation at a time

                # Update chart configuration
                charts[idx]['x_col'] = new_x_col
                charts[idx]['y_col'] = new_y_col
                charts[idx]['chart_type'] = new_chart_type
                charts[idx]['period'] = new_period
                charts[idx]['agg'] = new_agg
                charts[idx]['calc'] = new_calc
                charts[idx]['x_calc'] = new_x_calc
                charts[idx]['sort'] = new_sort
                charts[idx]['topn'] = new_topn
                charts[idx]['topn_n'] = new_topn_n

                # Check if the chart parameters actually changed
                chart_changed = (
                    old_chart.get('x_col', '') != new_x_col
                    or old_chart.get('y_col', '') != new_y_col
                    or old_chart.get('chart_type', 'auto') != new_chart_type
                    or old_chart.get('period', 'none') != new_period
                    or old_chart.get('agg', 'sum') != new_agg
                    or old_chart.get('calc') != new_calc
                    or old_chart.get('x_calc') != new_x_calc
                    or _rank_signature(old_chart.get('sort', 'default'), old_chart.get('topn', 'all'),
                                       old_chart.get('topn_n', 10)) != _rank_signature(new_sort, new_topn, new_topn_n)
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
        elif action in ('add_filter', 'remove_filter', 'clear_filters'):
            anchor = 'filters'
            old_filters = request.session.get('filters', [])
            new_filters = list(old_filters)
            if action == 'clear_filters':
                new_filters = []
            elif action == 'remove_filter':
                fidx = int(request.POST.get('filter_index', -1))
                if 0 <= fidx < len(new_filters):
                    new_filters.pop(fidx)
            else:
                try:
                    df_check = read_dataset(file_path)
                except Exception:
                    df_check = None
                col = request.POST.get('filter_col', '')
                if df_check is not None and col in df_check.columns:
                    is_date = column_kind(df_check[col]) == 'date'
                    spec = normalize_filter(df_check, {
                        'col': col,
                        'min': request.POST.get('filter_min_date' if is_date else 'filter_min_num'),
                        'max': request.POST.get('filter_max_date' if is_date else 'filter_max_num'),
                        'values': request.POST.getlist('filter_values'),
                        'text': request.POST.get('filter_text'),
                    })
                    if spec:
                        # one filter per column: adding again replaces it
                        new_filters = [f for f in new_filters if f.get('col') != col] + [spec]
                        new_filters = new_filters[-MAX_FILTERS:]
            if new_filters != old_filters:
                request.session['filters'] = new_filters
                # saved insights describe the old rows, so let them regenerate
                overrides = {}
                for pivot_entry in pivots:
                    pivot_entry['insight'] = None
                    pivot_entry['recommendation'] = None
                reset_generated = True
        elif action in ('add_kpi', 'remove_kpi', 'reset_kpis'):
            anchor = 'kpis'
            try:
                df_check = read_dataset(file_path)
            except Exception:
                df_check = None
            if df_check is None:
                if action == 'reset_kpis':
                    request.session.pop('kpis', None)
            else:
                stored_specs = request.session.get('kpis')
                old_specs = list(default_kpi_specs(df_check) if stored_specs is None else stored_specs)
                specs = list(old_specs)
                if action == 'reset_kpis':
                    request.session.pop('kpis', None)
                    specs = list(default_kpi_specs(df_check))
                else:
                    if action == 'remove_kpi':
                        try:
                            kidx = int(request.POST.get('kpi_index', -1))
                        except (TypeError, ValueError):
                            kidx = -1
                        if 0 <= kidx < len(specs):
                            specs.pop(kidx)
                    else:
                        spec = normalize_kpi(df_check, {'col': request.POST.get('kpi_col', ''),
                                                        'agg': request.POST.get('kpi_agg', '')})
                        if spec and spec not in specs and len(specs) < MAX_KPIS:
                            specs.append(spec)
                    request.session['kpis'] = specs
                if specs != old_specs:
                    # The written summary quotes the Key Metrics, so a saved copy is now out of date
                    summary_override = None
        elif action in ('update_pivot', 'reset_pivot'):
            pidx = _to_int(request.POST.get('pivot_index'))
            anchor = f'pivot-{pidx}'
            if 0 <= pidx < len(pivots):
                if action == 'reset_pivot':
                    # back to the default settings, so the saved notes no longer match
                    pivots[pidx] = {'cfg': None, 'insight': None, 'recommendation': None}
                else:
                    try:
                        df_check = read_dataset(file_path)
                    except Exception:
                        df_check = None
                    if df_check is not None:
                        old_cfg = normalize_pivot_config(df_check, pivots[pidx].get('cfg'))
                        new_cfg = normalize_pivot_config(df_check, _read_pivot_post(request.POST))
                        pivots[pidx]['cfg'] = new_cfg
                        if new_cfg != old_cfg:
                            # Same rule as the charts: a changed pivot gets fresh auto-generated notes.
                            pivots[pidx]['insight'] = None
                            pivots[pidx]['recommendation'] = None
        elif action in ('save_pivot_text', 'reset_pivot_text'):
            pidx = _to_int(request.POST.get('pivot_index'))
            anchor = f'pivot-notes-{pidx}'
            if 0 <= pidx < len(pivots):
                if action == 'reset_pivot_text':
                    pivots[pidx]['insight'] = None
                    pivots[pidx]['recommendation'] = None
                else:
                    # An empty box falls back to the auto-generated text.
                    pivots[pidx]['insight'] = request.POST.get('pivot_insight_text', '').strip() or None
                    pivots[pidx]['recommendation'] = request.POST.get('pivot_recommendation_text', '').strip() or None

        # A saved summary/recommendation list describes the old charts,
        # so discard it and let the auto-generated version take over.
        if reset_generated:
            summary_override = None
            recommendations_override = None

        request.session['charts'] = charts
        request.session['insights_overrides'] = overrides
        request.session['summary_override'] = summary_override
        request.session['recommendations_override'] = recommendations_override
        request.session['pivots'] = pivots
        request.session['layout'] = _normalize_layout(layout, len(charts), len(pivots))
        for old_key in ('pivot', 'pivot_insight_override', 'pivot_recommendation_override'):
            request.session.pop(old_key, None)
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
    extras = {}

    try:
        df_full = read_dataset(file_path)
        rows, cols = df_full.shape
        column_names = df_full.columns.tolist()
        num_cols = df_full.select_dtypes(include='number').columns.tolist()
        date_cols = [c for c in column_names if _looks_like_date(df_full[c])]

        # Global filters: every chart, insight, KPI and pivot below uses the filtered rows
        df, applied_filters = apply_filters(df_full, request.session.get('filters', []))

        preview_html = df.head(10).to_html(classes='preview-table', index=False, escape=True)
        overrides = request.session.get('insights_overrides', {})
        charts_config = request.session.get('charts', [])
        currency = request.session.get('currency_symbols', {})

        for idx, config in enumerate(charts_config):
            x_col, y_col, _, _ = _resolve_columns(df, config.get('x_col', ''), config.get('y_col', ''))
            chart_type = config.get('chart_type', 'auto')
            period = config.get('period', 'none')
            agg = config.get('agg', 'sum')

            if df.empty:
                rendered_charts.append(_empty_chart_entry(idx, x_col, y_col, config))
                continue

            # Apply a calculated X/Y value and/or date grouping if requested
            plot = prepare_plot_data(df, x_col, y_col, config)
            plot_df, grouping_note, draw_type = plot.df, plot.note, plot.draw_type
            if plot_df is df:
                draw_type = chart_type

            plt.clf()
            fig, ax = plt.subplots(figsize=(10, 4.5))
            active_chart_type = draw_chart(ax, fig, plot_df, plot.x, plot.y, draw_type, num_cols,
                                           keep_order=bool(plot.ranked),
                                           by_time=(plot.ranked == 'time'))
            if plot.y_spec:
                _label_calc_axis(ax, active_chart_type, plot.y)
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

            # A grouped "count of records" is a row count, not money
            chart_currency = dict(currency)
            if plot_df is not df and plot.x_spec is None and plot.y_spec is None and plot.ranked in ('', 'time') and agg == 'count':
                chart_currency.pop(y_col, None)

            insights_text = overrides.get(str(idx))
            if insights_text is None:
                if plot.y_spec:
                    insights_text = generate_calc_insights(plot_df, plot.x, plot.y, active_chart_type, plot.y_spec)
                else:
                    insights_text = generate_insights(plot_df, plot.x, plot.y, active_chart_type, num_cols, chart_currency,
                                                      measure=agg if plot.ranked == 'time' else 'sum')
                insights_text += _ranked_insight_note(plot)

            rendered_charts.append({
                'index': idx,
                'x_col': x_col,
                'y_col': y_col,
                'chart_type': chart_type,
                'period': period,
                'agg': agg,
                'calc': config.get('calc'),
                'calc_label': plot.y if plot.y_spec else '',
                'x_calc': config.get('x_calc'),
                'x_calc_label': plot.x if plot.x_spec else '',
                'sort': config.get('sort', 'default'),
                'topn': config.get('topn', 'all'),
                'topn_n': config.get('topn_n', 10),
                'display_title': _chart_display_title(x_col, y_col, plot, active_chart_type),
                'active_chart_type': active_chart_type,
                'type_label': CHART_TYPE_LABELS.get(active_chart_type, str(active_chart_type).replace('_', ' ').title()),
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

    # Charts follow the order chosen on the page (and are numbered in that order)
    rendered_charts = _order_charts(request, rendered_charts)

    # Built first because the written summary quotes the Key Metrics
    extras = _dashboard_extras(request, df_full, df, applied_filters, currency)
    blocks = _build_blocks(request, rendered_charts, extras['pivots'])

    if df.empty:
        overall_summary = "No rows match the current filters, so there is nothing to summarize."
        recommendations = ["Adjust or clear the filters to bring rows back."]
    else:
        overall_summary = summary_override or generate_overall_summary(
            df,
            rendered_charts,
            extras.get('kpis')
        )

        recommendations = recommendations_override or generate_recommendations(
            df,
            rendered_charts,
            currency
        )
    context = {
        'file_name': file_name,
        'cleaning_actions': request.session.get('cleaning_actions'),
        # one-time message after a saved progress file was restored
        'progress_notice': request.session.pop('progress_notice', None),
        'max_progress_mb': MAX_PROGRESS_MB,
        'max_progress_bytes': MAX_PROGRESS_BYTES,
        'rendered_charts': rendered_charts,
        'blocks': blocks,
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
        'has_custom_text': (
            bool(request.session.get('insights_overrides'))
            or bool(summary_override)
            or bool(recommendations_override)
            or bool(extras.get('pivot_custom_text'))
        ),
    }
    context.update(extras)
    return render(request, 'cascade_app/visualize.html', context)
def _kpi_summary_sentence(kpis):
    """One sentence quoting the Key Metric cards, in the order they are shown."""
    parts = []
    for kpi in kpis or []:
        text = f"{kpi.get('label', '')} {kpi.get('value', '')}".strip()
        if kpi.get('sub'):
            text += f" ({kpi['sub']})"
        if text:
            parts.append(text)
    if not parts:
        return ''
    return "Key metrics: " + ", ".join(parts) + "."


def generate_overall_summary(df, rendered_charts, kpis=None):
    """
    Creates an overall summary from the Key Metrics, the generated chart
    insights and the actual dataset.
    """

    insights = [
        chart.get('insights_text', '').strip()
        for chart in rendered_charts
        if chart.get('insights_text', '').strip()
    ]
    kpi_sentence = _kpi_summary_sentence(kpis)

    if not insights and not kpi_sentence:
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

    # Key Metrics cards
    if kpi_sentence:
        summary_parts.append(kpi_sentence)

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
    number = chart.get('number') or chart.get('index', 0) + 1
    x_col = chart.get('x_calc_label') or chart.get('x_col', '')
    y_col = chart.get('calc_label') or chart.get('y_col', '')

    if active_type == 'heatmap':
        return f"Chart {number} - {label}"
    if active_type in ('histogram', 'box') and (x_col == y_col or chart.get('x_calc_label')):
        return f"Chart {number} - {label} ({x_col})"
    return f"Chart {number} - {label} ({y_col} by {x_col})"


def _recommend_for_chart(df, chart, active_type, num_cols, currency=None):
    """
    Recommendations that depend on what this specific chart shows.
    The numbers mirror what draw_chart() plots (e.g. bars use totals).
    """
    recs = []
    x_col = chart.get('x_col')
    y_col = chart.get('y_col')

    # If the chart is grouped by period or uses a calculated Y value,
    # advise on what is actually plotted.
    currency = dict(currency or {})
    original_df = df
    plot = None
    try:
        plot = prepare_plot_data(df, x_col, y_col, chart)
        df, x_col, y_col = plot.df, plot.x, plot.y
        # A "count of records" is a number of rows, not an amount of money.
        if (df is not original_df and plot.x_spec is None
                and plot.y_spec is None and plot.ranked in ('', 'time') and chart.get('agg') == 'count'):
            currency.pop(y_col, None)
    except Exception:
        pass

    if plot is not None and plot.y_spec:
        return _recommend_for_calc(df, x_col, y_col, active_type, plot.y_spec)

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
                f"describe '{x_col}' with the median ({_money(median, x_col, currency, ',.2f')}) rather than "
                f"the mean ({_money(mean, x_col, currency, ',.2f')}) and check the extreme {side} values "
                f"for errors or special cases."
            )
        else:
            recs.append(
                f"The distribution is fairly symmetric, so the mean "
                f"({_money(mean, x_col, currency, ',.2f')}) is a reliable summary of typical '{x_col}' "
                f"values. Focus on whether the spread (standard deviation "
                f"{_money(series.std(), x_col, currency, ',.2f')}) is acceptable for your goals."
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
        # Bars grouped by day/week/month/... : one bar per period, in time order
        time_grouped = plot is not None and plot.ranked == 'time' and df is not original_df
        if y_numeric and x_col != y_col:
            grouped = df.groupby(x_col)[y_col].sum()
            counts = None if time_grouped else df.groupby(x_col)[y_col].count()
            word = {'mean': 'average', 'max': 'maximum', 'min': 'minimum'}.get(chart.get('agg'), 'total')
            measure = "number of records" if (time_grouped and chart.get('agg') == 'count') else f"{word} '{y_col}'"
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

        if time_grouped:
            recs.append(
                f"'{top}' is the highest period and '{low}' the lowest in {measure}. "
                f"Check what changed between them (seasonality, promotions, supply "
                f"or staffing) before planning around either."
            )
        elif top_share is not None and top_share >= 40:
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
            if n_groups > limit and not time_grouped:
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
            x_dates = _to_datetime(clean[x_col])
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
        peak_s = _money(values[peak_pos], y_col, currency, ',.2f')
        low_s = _money(values[low_pos], y_col, currency, ',.2f')
        first_s = _money(first, y_col, currency, ',.2f')
        last_s = _money(last, y_col, currency, ',.2f')

        if first != 0:
            change = (last - first) / abs(first) * 100
            if change >= 10:
                recs.append(
                    f"'{y_col}' rose from {first_s} to {last_s} "
                    f"({change:+.1f}%). Identify what is driving the growth and "
                    f"whether it can be sustained."
                )
            elif change <= -10:
                recs.append(
                    f"'{y_col}' fell from {first_s} to {last_s} "
                    f"({change:+.1f}%). Determine the cause of the decline and "
                    f"whether corrective action is needed."
                )
            else:
                recs.append(
                    f"'{y_col}' is fairly stable overall ({change:+.1f}% "
                    f"from start to end). Focus on the peak at '{peak_x}' "
                    f"({peak_s}) and the low at '{low_x}' "
                    f"({low_s}) to find what caused them."
                )
        else:
            recs.append(
                f"'{y_col}' peaks at '{peak_x}' ({peak_s}) "
                f"and is lowest at '{low_x}' ({low_s}). "
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
                    f"({_money(medians.max(), y_col, currency, ',.2f')}) and '{lo}' the lowest "
                    f"({_money(medians.min(), y_col, currency, ',.2f')}). Investigate why these groups differ."
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


# "Chart 1 - Bar Chart (Profit by Region): advice..."  ->  chart, type, subject, advice
_REC_CHART_PATTERN = re.compile(r'^Chart (\d+) - (.+?)(?: \((.+?)\))?: (.*)$', re.DOTALL)


def _recommendation_items(recommendations):
    """Split recommendation sentences into a heading and the advice itself.

    Used by the PDF so the chart a recommendation refers to can be shown as a
    clean heading (e.g. 'Chart 1 · Bar Chart') with the advice underneath.
    Anything that doesn't start with a chart reference gets a plain label.
    """
    items = []
    for text in recommendations or []:
        text = str(text).strip()
        if not text:
            continue
        match = _REC_CHART_PATTERN.match(text)
        if match:
            items.append({
                'ref': f"Chart {match.group(1)} · {match.group(2).strip()}",
                'subject': (match.group(3) or '').strip(),
                'body': match.group(4).strip(),
                'chart_number': int(match.group(1)),
            })
        else:
            label = 'Data Quality' if text.lower().startswith('resolve missing data') else 'General'
            items.append({'ref': label, 'subject': '', 'body': text, 'chart_number': None})
    return items


def generate_recommendations(df, rendered_charts, currency=None):
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

        cols = num_cols if active_type == 'heatmap' else _chart_source_cols(chart)
        for col in cols:
            if col in df.columns and col not in charted_cols:
                charted_cols.append(col)

        try:
            advice = _recommend_for_chart(df, chart, active_type, num_cols, currency)
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
        df_full = read_dataset(file_path)
        df, applied_filters = apply_filters(df_full, request.session.get('filters', []))
        if df.empty:
            return redirect('visualize_data')

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

        currency = request.session.get('currency_symbols', {})

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

            # Apply a calculated X/Y value and/or date grouping if requested
            plot = prepare_plot_data(df, x_col, y_col, config)
            plot_df, grouping_note, draw_type = plot.df, plot.note, plot.draw_type
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
                plot.x,
                plot.y,
                draw_type,
                num_cols,
                keep_order=bool(plot.ranked),
                by_time=(plot.ranked == 'time')
            )

            if plot.y_spec:
                _label_calc_axis(ax, active_chart_type, plot.y)

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
            # A grouped "count of records" is a row count, not money
            chart_currency = dict(currency)
            if plot_df is not df and plot.x_spec is None and plot.y_spec is None and plot.ranked in ('', 'time') and agg == 'count':
                chart_currency.pop(y_col, None)

            insights_text = overrides.get(
                str(idx)
            )

            if insights_text is None:
                if plot.y_spec:
                    insights_text = generate_calc_insights(
                        plot_df,
                        plot.x,
                        plot.y,
                        active_chart_type,
                        plot.y_spec
                    )
                else:
                    insights_text = generate_insights(
                        plot_df,
                        plot.x,
                        plot.y,
                        active_chart_type,
                        num_cols,
                        chart_currency,
                        measure=agg if plot.ranked == 'time' else 'sum'
                    )
                insights_text += _ranked_insight_note(plot)

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
                'calc': config.get('calc'),
                'calc_label': plot.y if plot.y_spec else '',
                'x_calc': config.get('x_calc'),
                'x_calc_label': plot.x if plot.x_spec else '',
                'sort': config.get('sort', 'default'),
                'topn': config.get('topn', 'all'),
                'topn_n': config.get('topn_n', 10),
                'display_title': _chart_display_title(x_col, y_col, plot, active_chart_type),
                'active_chart_type': active_chart_type,
                'image_uri': uri,
                'insights_text': insights_text
            })


        # ---------------------------------------------------------
        # OVERALL SUMMARY
        # ---------------------------------------------------------
        summary_override = request.session.get('summary_override')
        recommendations_override = request.session.get('recommendations_override')

        # Same order and numbering as the page
        rendered_charts = _order_charts(request, rendered_charts)
        extras = _dashboard_extras(request, df_full, df, applied_filters, currency)
        blocks = _build_blocks(request, rendered_charts, extras['pivots'], skip_broken_pivots=True)

        overall_summary = summary_override or generate_overall_summary(
            df,
            rendered_charts,
            extras.get('kpis')
        )

        # ---------------------------------------------------------
        # RECOMMENDATIONS
        # ---------------------------------------------------------
        recommendations = recommendations_override or generate_recommendations(
            df,
            rendered_charts,
            currency
        )

        # Each chart also carries its own recommendations, shown right under it;
        # the full list still closes the report.
        recommendation_items = _recommendation_items(recommendations)
        for chart in rendered_charts:
            chart['recommendation_items'] = [
                item for item in recommendation_items if item.get('chart_number') == chart.get('number')
            ]

        # ---------------------------------------------------------
        # PDF CONTEXT
        # ---------------------------------------------------------
        author_name = (
            request.user.get_full_name()
            or request.user.username
        )

        context = {
            'file_name': file_name,

            # What the user chose during cleaning (missing-value decisions)
            'cleaning_actions': request.session.get('cleaning_actions'),

            'charts': rendered_charts,

            'author_name': author_name,

            # Dataset information
            'rows': rows,
            'cols': cols,

            # Overall analysis
            'overall_summary': overall_summary,
            'recommendations': recommendations,
            'recommendation_items': recommendation_items,
            'blocks': blocks,
        }
        # Same filters, KPI cards and pivot table as the dashboard
        # (available in pdf_template.html as kpis, pivot and filter_chips)
        context.update(extras)

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
    if ext in JSON_EXTENSIONS:
        return _read_json_dataset(file_path)
    return pd.read_csv(file_path)


def _read_json_dataset(file_path):
    """Load a JSON file as a flat table (one row per record).

    Accepted shapes:
      - a list of objects:             [{"a": 1}, {"a": 2}]
      - JSON Lines (.jsonl/.ndjson):   one object per line
      - an object wrapping a list:     {"data": [{"a": 1}, ...]}
      - an object of equal columns:    {"a": [1, 2], "b": [3, 4]}
      - a single object (one row):     {"a": 1, "b": 2}
    Nested objects are flattened into columns named parent_child. Lists left
    inside a cell are stored as JSON text so duplicate checks still work.
    """
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        text = f.read()

    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        # Not one JSON document, so try JSON Lines.
        try:
            data = [json.loads(line) for line in text.splitlines() if line.strip()]
        except json.JSONDecodeError:
            raise ValueError("That file is not valid JSON.")

    records = None
    if isinstance(data, list):
        records = data
    elif isinstance(data, dict):
        # Look for a list of objects inside the wrapper, e.g. {"results": [...]}.
        wrapped = [v for v in data.values()
                   if isinstance(v, list) and v and all(isinstance(r, dict) for r in v)]
        if wrapped:
            records = max(wrapped, key=len)
        elif data and all(isinstance(v, list) for v in data.values()) \
                and len({len(v) for v in data.values()}) == 1:
            return _flatten_json_cells(pd.DataFrame(data))
        else:
            records = [data]

    if not records or not all(isinstance(r, dict) for r in records):
        raise ValueError(
            "Could not turn that JSON into a table. Use a list of records, "
            "for example [{\"region\": \"North\", \"sales\": 100}, ...]."
        )

    return _flatten_json_cells(pd.json_normalize(records, sep='_'))


def _flatten_json_cells(df):
    """Store any list/dict left in a cell as JSON text so rows stay hashable."""
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(
                lambda v: json.dumps(v) if isinstance(v, (list, dict)) else v
            )
    return df

# ---------------------------------------------------------------------
# DATE RECOGNITION
# ---------------------------------------------------------------------
# Only text that really is a date counts as one. Bare month names ("March"),
# bare years ("2024"), quarters ("2024-Q1"), weeks and codes ("T-001") are
# labels, not dates: they are left alone so they stay usable as categories
# (and keep their checkbox filter). Month-level values ("Jan 2024", "2024-01",
# "Jan-24") are dates too, but they are never turned into a made-up day.

_MONTH_LOOKUP = {}
for _num, _name in enumerate(
        ['january', 'february', 'march', 'april', 'may', 'june', 'july',
         'august', 'september', 'october', 'november', 'december'], start=1):
    _MONTH_LOOKUP[_name] = _num
    _MONTH_LOOKUP[_name[:3]] = _num
_MONTH_LOOKUP['sept'] = 9


def _month_number(word):
    key = word.lower().rstrip('.')
    return _MONTH_LOOKUP.get(key) or _MONTH_LOOKUP.get(key[:3])

_MON = (r"(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|"
        r"aug(?:ust)?|sept?(?:ember)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\.?")
# optional time of day (24 h or 12 h with AM/PM) and time-zone tail
_TIME = (r"(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2})(?:\.\d+)?)?\s*([ap]m)?"
         r"\s*(?:Z|[+-]\d{2}:?\d{2})?)?")

_DATE_FORMS = [
    # (name, compiled pattern). Group order is documented per name below.
    ('ymd',   re.compile(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})" + _TIME + r"$", re.I)),
    ('abyyyy', re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})" + _TIME + r"$", re.I)),
    ('abyy',  re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{2})" + _TIME + r"$", re.I)),
    ('dmony', re.compile(r"^(\d{1,2})(?:st|nd|rd|th)?[ \-/.]+" + _MON + r"[ \-/.,]+(\d{4}|\d{2})" + _TIME + r"$", re.I)),
    ('mondy', re.compile(r"^" + _MON + r"[ \-/]+(\d{1,2})(?:st|nd|rd|th)?,?[ \-/]+(\d{4})" + _TIME + r"$", re.I)),
    ('ym',    re.compile(r"^(\d{4})[-/](\d{1,2})$")),
    ('my',    re.compile(r"^(\d{1,2})[-/](\d{4})$")),
    ('mony4', re.compile(r"^" + _MON + r"[ \-/.,']*(\d{4})$", re.I)),
    ('mony2', re.compile(r"^" + _MON + r"[\-']\s?(\d{2})$", re.I)),
    ('ymon',  re.compile(r"^(\d{4})[ \-/.,]*" + _MON + r"$", re.I)),
]
_NUMERIC_AB = re.compile(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})\b|^(\d{1,2})/(\d{1,2})/(\d{2})\b")
_ISO_FULL = re.compile(r"^\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?)?$")
MIN_DATE_YEAR, MAX_DATE_YEAR = 1000, 2999


def _two_digit_year(yy):
    return 2000 + yy if yy < 69 else 1900 + yy


def _build_datetime(year, month, day, tail):
    """datetime from parts plus the optional (hour, minute, second, am/pm)
    tail of a pattern, or None when any part is impossible."""
    import datetime
    if not (MIN_DATE_YEAR <= year <= MAX_DATE_YEAR):
        return None
    hour = minute = second = 0
    if tail and tail[0] is not None:
        hour, minute = int(tail[0]), int(tail[1])
        second = int(tail[2]) if tail[2] else 0
        meridiem = (tail[3] or '').lower()
        if meridiem:
            if not 1 <= hour <= 12:
                return None
            hour = hour % 12 + (12 if meridiem == 'pm' else 0)
    try:
        return datetime.datetime(year, month, day, hour, minute, second)
    except ValueError:
        return None


def _parse_date_text(text, dayfirst=None):
    """Read one piece of text as a date.

    Returns (datetime, 'full' | 'month'), or (None, None) when it is not a
    real date. dayfirst: True = dd/mm/yyyy, False = mm/dd/yyyy, None = decide
    from the value itself (a number above 12 can only be the day) and read
    month first when it could be either."""
    text = str(text).strip()
    if not text or len(text) > 40:
        return None, None
    stamp, form = _read_date_form(text, dayfirst)
    return (stamp, form) if stamp is not None else (None, None)


def _read_date_form(text, dayfirst):
    for name, pattern in _DATE_FORMS:
        m = pattern.match(text)
        if not m:
            continue
        g = m.groups()
        if name == 'ymd':
            return _build_datetime(int(g[0]), int(g[1]), int(g[2]), g[3:]), 'full'
        if name in ('abyyyy', 'abyy'):
            a, b = int(g[0]), int(g[1])
            year = int(g[2]) if name == 'abyyyy' else _two_digit_year(int(g[2]))
            if dayfirst is True:
                day, month = a, b
            elif dayfirst is False:
                month, day = a, b
            elif a > 12 >= b:
                day, month = a, b
            else:
                month, day = a, b
            return _build_datetime(year, month, day, g[3:]), 'full'
        if name == 'dmony':
            year = int(g[2]) if len(g[2]) == 4 else _two_digit_year(int(g[2]))
            return _build_datetime(year, _month_number(g[1]), int(g[0]), g[3:]), 'full'
        if name == 'mondy':
            month = _month_number(g[0])
            return _build_datetime(int(g[2]), month, int(g[1]), g[3:]), 'full'
        # month-level shapes: the 1st of the month stands in for "the month"
        if name == 'ym':
            year, month = int(g[0]), int(g[1])
        elif name == 'my':
            year, month = int(g[1]), int(g[0])
        elif name == 'ymon':
            year, month = int(g[0]), _month_number(g[1])
        else:  # mony4 / mony2
            month = _month_number(g[0])
            year = int(g[1]) if name == 'mony4' else _two_digit_year(int(g[1]))
        stamp = _build_datetime(year, month, 1, None)
        return (stamp, 'month') if stamp else (None, None)
    return None, None


def _infer_dayfirst(texts):
    """Look at every dd/mm-or-mm/dd value of a column and decide the order.
    Returns (dayfirst, ambiguous): dayfirst is True, False, or None when the
    column mixes both orders; ambiguous is True when nothing in the column
    settles it (every day and month is 12 or lower), so month/day is assumed."""
    first_big = second_big = differs = False
    for text in texts:
        m = _NUMERIC_AB.match(str(text).strip())
        if not m:
            continue
        a, b = (int(x) for x in (m.group(1), m.group(2))) if m.group(1) else (int(m.group(4)), int(m.group(5)))
        if a > 12 and b <= 12:
            first_big = True
        elif b > 12 and a <= 12:
            second_big = True
        elif a <= 12 and b <= 12 and a != b:
            differs = True
    if first_big and second_big:
        return None, False
    if first_big:
        return True, False
    if second_big:
        return False, False
    return False, differs


def _looks_like_date(series, sample_size=200):
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    if series.dtype.kind in 'bifc':
        return False
    sample = series.dropna().astype(str).str.strip()
    sample = sample[sample != ''].head(sample_size)
    if sample.empty:
        return False
    try:
        dayfirst, _ = _infer_dayfirst(sample)
        found = sum(1 for text in sample if _parse_date_text(text, dayfirst)[0] is not None)
        return found / len(sample) > 0.8
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

    dates = _to_datetime(df[x_col])
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


def normalize_calc(df, raw):
    """Check a saved calculation against the dataset.

    Returns a clean dict, or None when the calculation is off or can't be
    used (unknown column, non-numeric column, bad number, divide by zero).
    """
    if not isinstance(raw, dict) or not raw.get('on'):
        return None

    numeric = set(df.select_dtypes(include='number').columns)
    a = raw.get('a')
    op = raw.get('op')
    if a not in numeric or op not in CALC_OPS:
        return None

    b_kind = raw.get('b_kind')
    b = str(raw.get('b', '')).strip()
    if b_kind == 'col':
        if b not in numeric:
            return None
    elif b_kind == 'num':
        try:
            value = float(b)
        except ValueError:
            return None
        if value != value or value in (float('inf'), float('-inf')):
            return None
        if op == '/' and value == 0:
            return None
    else:
        return None

    mode = raw.get('mode') if raw.get('mode') in CALC_MODES else 'total'
    return {
        'a': a, 'op': op, 'b_kind': b_kind, 'b': b,
        'pct': bool(raw.get('pct')), 'mode': mode,
    }


def calc_label(spec):
    """Readable name of the calculated value, e.g. 'Profit ÷ Net Revenue × 100'."""
    b = spec['b'] if spec['b_kind'] == 'col' else f"{float(spec['b']):g}"
    label = f"{spec['a']} {CALC_OPS[spec['op']]} {b}"
    if spec['pct']:
        label += " × 100"
    return label


def _calc_is_ratio_of_totals(spec):
    return spec['mode'] == 'total' and spec['op'] == '/' and spec['b_kind'] == 'col'


def calc_is_additive(spec):
    """True when group results can be added up (so a pie of shares makes sense)."""
    return spec['mode'] == 'total' and not (spec['op'] == '/' and spec['b_kind'] == 'col')


def calc_note(spec):
    """Short description of how the rows were combined (shown in the title)."""
    if _calc_is_ratio_of_totals(spec):
        return "ratio of totals"
    if spec['mode'] == 'mean':
        return "average of each row"
    return "total of each row"


def prepare_calculation(df, x_col, config):
    """Build the calculated Y value, one result per X group.

    Returns (plot_df, label, note, draw_type), or None when the chart has no
    (usable) calculation. plot_df has two columns: [x_col, label].
    """
    spec = normalize_calc(df, config.get('calc'))
    if spec is None or x_col not in df.columns:
        return None

    chart_type = config.get('chart_type', 'auto')
    try:
        resolved = detect_chart_type(df, x_col, spec['a']) if chart_type == 'auto' else chart_type
    except Exception:
        return None
    if resolved not in CALC_CHART_TYPES:
        return None

    label = calc_label(spec)
    note = calc_note(spec)
    draw_type = resolved
    if resolved == 'pie' and not calc_is_additive(spec):
        draw_type = 'bar'
        note += "; drawn as bars because ratios and averages can't be pie shares"

    a = pd.to_numeric(df[spec['a']], errors='coerce')
    if spec['b_kind'] == 'col':
        b = pd.to_numeric(df[spec['b']], errors='coerce')
    else:
        b = pd.Series(float(spec['b']), index=df.index)

    period = config.get('period', 'none')
    use_period = (
        period in TIME_PERIODS and period != 'none'
        and resolved in TIME_GROUPABLE_TYPES
        and _looks_like_date(df[x_col])
    )
    if use_period:
        freq = TIME_PERIODS[period][1]
        keys = _to_datetime(df[x_col]).dt.to_period(freq)
    else:
        keys = df[x_col]

    # Only rows where every part is present, so numerator and denominator
    # always cover exactly the same rows.
    work = pd.DataFrame({'key': keys, 'a': a, 'b': b}).dropna()
    if work.empty:
        return None

    if _calc_is_ratio_of_totals(spec):
        sums = work.groupby('key')[['a', 'b']].sum()
        result = sums['a'] / sums['b'].where(sums['b'] != 0)
    else:
        if spec['op'] == '+':
            row = work['a'] + work['b']
        elif spec['op'] == '-':
            row = work['a'] - work['b']
        elif spec['op'] == '*':
            row = work['a'] * work['b']
        else:
            row = work['a'] / work['b'].where(work['b'] != 0)
        grouped = row.groupby(work['key'])
        result = grouped.sum(min_count=1) if spec['mode'] == 'total' else grouped.mean()

    if spec['pct']:
        result = result * 100
    result = result.replace([float('inf'), float('-inf')], float('nan'))

    if use_period:
        # Keep empty periods so the timeline has no silent gaps
        full_range = pd.period_range(result.index.min(), result.index.max(), freq=freq)
        result = result.reindex(full_range)
        if spec['mode'] == 'total' and not _calc_is_ratio_of_totals(spec):
            result = result.fillna(0)
        x_values = [_period_label(p, freq) for p in result.index]
    else:
        result = result.dropna()
        if result.empty:
            return None
        x_values = result.index.tolist()

    plot_df = pd.DataFrame({x_col: x_values, label: result.to_numpy(dtype=float)})
    return plot_df, label, note, draw_type, spec


def _read_calc_post(post, prefix):
    """Read one calculation ('calc' = Y, 'xcalc' = X) from the submitted form."""
    if post.get(f'{prefix}_on') != '1':
        return None
    b_kind = post.get(f'{prefix}_b_kind', 'col')
    return {
        'on': True,
        'a': post.get(f'{prefix}_a', ''),
        'op': post.get(f'{prefix}_op', '/'),
        'b_kind': b_kind,
        'b': post.get(f'{prefix}_b_col' if b_kind == 'col' else f'{prefix}_b_num', '').strip(),
        'pct': post.get(f'{prefix}_pct') == '1',
        'mode': post.get(f'{prefix}_mode', 'total'),
    }


def prepare_x_calculation(df, x_col, y_col, config):
    """Build a calculated X value, worked out for every row.

    Used by scatter, histogram and box charts. Returns PlotData, or None when
    there is no (usable) X calculation for this chart.
    """
    spec = normalize_calc(df, config.get('x_calc'))
    if spec is None:
        return None

    label = calc_label(spec)
    if label in df.columns:
        return None

    chart_type = config.get('chart_type', 'auto')
    if chart_type == 'auto':
        # A numeric X against a numeric Y is a scatter plot
        if y_col not in df.columns or df[y_col].dtype.kind not in 'bifc':
            return None
        resolved = 'scatter'
    else:
        resolved = chart_type
    if resolved not in X_CALC_CHART_TYPES:
        return None

    a = pd.to_numeric(df[spec['a']], errors='coerce')
    if spec['b_kind'] == 'col':
        b = pd.to_numeric(df[spec['b']], errors='coerce')
    else:
        b = pd.Series(float(spec['b']), index=df.index)

    if spec['op'] == '+':
        values = a + b
    elif spec['op'] == '-':
        values = a - b
    elif spec['op'] == '*':
        values = a * b
    else:
        values = a / b.where(b != 0)
    if spec['pct']:
        values = values * 100
    values = values.replace([float('inf'), float('-inf')], float('nan'))
    if values.notna().sum() == 0:
        return None

    plot_df = df.copy()
    plot_df[label] = values

    # Histogram and box plots describe the calculated value on its own
    plot_y = label if resolved in ('histogram', 'box') else y_col
    return PlotData(plot_df, label, plot_y, '', resolved, spec, None)


def prepare_plot_data(df, x_col, y_col, config):
    """Single entry point for "what is actually plotted".

    Applies, in order: a calculated X value, a calculated Y value, or date
    grouping. When nothing changes, `df` itself comes back untouched.
    """
    x_result = prepare_x_calculation(df, x_col, y_col, config)
    if x_result is not None:
        return x_result

    calculated = prepare_calculation(df, x_col, config)
    if calculated is not None:
        plot_df, label, note, draw_type, spec = calculated
        plot = PlotData(plot_df, x_col, label, note, draw_type, None, spec)
    else:
        plot_df, note, draw_type = prepare_time_grouping(df, x_col, y_col, config)
        plot = PlotData(plot_df, x_col, y_col, note, draw_type, None, None)

    # Sort / Top-Bottom N for bar, horizontal bar and pie charts
    plot = prepare_ranking(df, x_col, y_col, config, plot)

    # A chart grouped by day/week/month/... keeps time order, with no cut-off
    period = config.get('period', 'none')
    if (not plot.ranked and period in TIME_PERIODS and period != 'none'
            and plot.x_spec is None and plot.df is not df
            and plot.draw_type in TIME_GROUPABLE_TYPES
            and x_col in df.columns and _looks_like_date(df[x_col])):
        plot = plot._replace(ranked='time')
    return plot


# =====================================================================
# DASHBOARD TOOLS: Filters, KPI cards, Sort / Rank / Top-Bottom N, Pivot table
#
# Nothing in here knows any column name. Every feature looks at the column
# TYPES of whatever file was uploaded (number / date / text), so it works on
# any dataset.
# =====================================================================

RECORDS_COL = '__records__'      # pseudo-column meaning "count of records"
COUNT_LABEL = 'Count of records'

# --- Sort / Top-Bottom N for charts ---
SORT_OPTIONS = {
    'default': 'Default order',
    'value_desc': 'Value: high to low',
    'value_asc': 'Value: low to high',
    'label_asc': 'Label: A to Z',
    'label_desc': 'Label: Z to A',
}
SORT_NOTES = {
    'value_desc': 'high to low', 'value_asc': 'low to high',
    'label_asc': 'A to Z', 'label_desc': 'Z to A',
}
TOPN_MODES = {'all': 'Show all', 'top': 'Top N', 'bottom': 'Bottom N'}
RANKABLE_TYPES = ('bar', 'horizontal_bar', 'pie')   # one value per group
MAX_TOP_N = 50

# --- Filters ---
MAX_FILTERS = 12
MAX_FILTER_CHOICES = 200         # above this a text column gets a "contains" box

# --- KPI cards ---
MAX_KPIS = 12
KPI_AGGS = {
    'sum': 'Total', 'mean': 'Average', 'median': 'Median',
    'min': 'Minimum', 'max': 'Maximum',
    'count': 'Count (non-blank)', 'nunique': 'Distinct count', 'mode': 'Most common',
}

# --- Pivot table ---
PIVOT_AGGS = {
    'sum': 'Sum', 'mean': 'Average', 'median': 'Median', 'min': 'Minimum',
    'max': 'Maximum', 'count': 'Count', 'nunique': 'Distinct count',
}
PIVOT_SORTS = {'total': 'Total / value', 'label': 'Row label'}
MAX_PIVOT_COLS = 15
MAX_PIVOT_ROWS = 200


def _pretty(name):
    return str(name).replace('_', ' ').strip().title()


def _to_float(value):
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    if number != number or number in (float('inf'), float('-inf')):
        return None
    return number


def _plain_num(value):
    return f"{value:,.2f}".rstrip('0').rstrip('.')


def _to_datetime(series):
    """Dates of a column as real timestamps (NaT where a value is not a date).
    Reads the same shapes _looks_like_date accepts, so the two never disagree."""
    if pd.api.types.is_datetime64_any_dtype(series):
        try:
            return series.dt.tz_localize(None)
        except TypeError:
            return series
    if series.dtype.kind in 'bifc':
        return pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')

    present = series.dropna()
    if present.empty:
        return pd.Series(pd.NaT, index=series.index, dtype='datetime64[ns]')

    # Fast path: clean ISO text (what the cleaning step writes)
    as_text = present.astype(str).str.strip()
    if as_text.str.fullmatch(_ISO_FULL.pattern).all():
        try:
            parsed = pd.to_datetime(series.astype(str).str.strip().where(series.notna()),
                                    errors='coerce', format='ISO8601')
            return parsed
        except (TypeError, ValueError):
            pass

    uniques = pd.unique(present)
    texts = {u: str(u).strip() for u in uniques}
    dayfirst, _ = _infer_dayfirst(texts.values())
    stamps = {u: _parse_date_text(text, dayfirst)[0] for u, text in texts.items()}
    return pd.to_datetime(series.map(stamps), errors='coerce')


def _is_month_level(series):
    """True for a date column whose values are whole months (YYYY-MM)."""
    if series.dtype.kind in 'bifcM':
        return False
    present = series.dropna().astype(str)
    return bool(len(present)) and bool(present.str.fullmatch(r"\d{4}-\d{2}").all())


def column_kind(series):
    """'number', 'date' or 'text'. Decides which options a column is offered."""
    if series.dtype.kind == 'b':
        return 'text'
    if series.dtype.kind in 'ifc':
        return 'number'
    if _looks_like_date(series):
        return 'date'
    return 'text'


def _is_identifier_like(series, name):
    """True for columns that are labels in disguise (IDs, years, zip codes...),
    so they are not picked automatically as something to add up."""
    s = series.dropna()
    if len(s) < 3 or s.dtype.kind not in 'iuf':
        return False
    if (s % 1 != 0).any():
        return False
    tokens = set(re.split(r'[^a-z]+', str(name).lower()))
    if 'year' in tokens:
        return True
    if tokens & {'id', 'uid', 'no', 'code', 'index', 'zip', 'phone'} and s.nunique() / len(s) > 0.5:
        return True
    return bool(s.is_unique and (s.is_monotonic_increasing or s.is_monotonic_decreasing))


# ---------------------------------------------------------------------
# FILTERS
# ---------------------------------------------------------------------

def build_filter_meta(df):
    """Per-column info the page needs to draw the right filter inputs."""
    meta = {}
    for col in df.columns:
        s = df[col]
        kind = column_kind(s)
        info = {'kind': kind}
        if kind == 'number':
            clean = pd.to_numeric(s, errors='coerce').dropna()
            if len(clean):
                info['min'] = float(clean.min())
                info['max'] = float(clean.max())
        elif kind == 'date':
            dates = _to_datetime(s).dropna()
            if len(dates):
                info['min'] = dates.min().strftime('%Y-%m-%d')
                info['max'] = dates.max().strftime('%Y-%m-%d')
        else:
            if s.nunique(dropna=True) <= MAX_FILTER_CHOICES:
                choices = sorted({str(v) for v in s.dropna().unique()}, key=str.casefold)
                if s.isna().any():
                    choices.append('(blank)')
                info['choices'] = choices
        meta[str(col)] = info
    return meta


def normalize_filter(df, raw):
    """Check a saved filter against the dataset. Returns a clean dict, or None
    when it is empty or refers to a column that no longer exists."""
    if not isinstance(raw, dict) or raw.get('col') not in df.columns:
        return None
    col = raw['col']
    kind = column_kind(df[col])
    spec = {'col': col, 'kind': kind}

    if kind == 'number':
        lo, hi = _to_float(raw.get('min')), _to_float(raw.get('max'))
        if lo is None and hi is None:
            return None
        if lo is not None and hi is not None and lo > hi:
            lo, hi = hi, lo
        spec.update(min=lo, max=hi)
    elif kind == 'date':
        bounds = []
        for key in ('min', 'max'):
            text = str(raw.get(key) or '').strip()
            stamp = pd.to_datetime(text, errors='coerce') if text else pd.NaT
            bounds.append(None if pd.isna(stamp) else stamp.strftime('%Y-%m-%d'))
        lo, hi = bounds
        if lo is None and hi is None:
            return None
        if lo is not None and hi is not None and lo > hi:
            lo, hi = hi, lo
        spec.update(min=lo, max=hi)
    else:
        values = []
        for v in raw.get('values') or []:
            v = str(v)
            if v != '' and v not in values:
                values.append(v)
        text = str(raw.get('text') or '').strip()
        if values:
            spec['values'] = values
        elif text:
            spec['text'] = text
        else:
            return None
    return spec


def apply_filters(df, filters):
    """Apply every saved filter (AND). Returns (filtered_df, [(index, spec), ...])
    where index is the filter's position in the saved list (used by the remove buttons).
    When nothing is applied the original `df` object comes back untouched."""
    mask = pd.Series(True, index=df.index)
    applied = []

    for i, raw in enumerate(filters or []):
        spec = normalize_filter(df, raw)
        if spec is None:
            continue
        series = df[spec['col']]

        if spec['kind'] == 'number':
            values = pd.to_numeric(series, errors='coerce')
            if spec['min'] is not None:
                mask &= values >= spec['min']
            if spec['max'] is not None:
                mask &= values <= spec['max']
        elif spec['kind'] == 'date':
            days = _to_datetime(series).dt.normalize()
            if _is_month_level(series):
                # A month stands for the whole month, so compare months:
                # "2024-03" is inside 15 Mar - 20 Apr, not outside it.
                months = days.dt.to_period('M')
                if spec['min'] is not None:
                    mask &= months >= pd.Timestamp(spec['min']).to_period('M')
                if spec['max'] is not None:
                    mask &= months <= pd.Timestamp(spec['max']).to_period('M')
            else:
                if spec['min'] is not None:
                    mask &= days >= pd.Timestamp(spec['min'])
                if spec['max'] is not None:
                    mask &= days <= pd.Timestamp(spec['max'])
        else:
            as_text = series.astype(str).where(series.notna(), '(blank)')
            if 'values' in spec:
                mask &= as_text.isin(spec['values'])
            else:
                mask &= as_text.str.contains(spec['text'], case=False, regex=False, na=False)
        applied.append((i, spec))

    if not applied:
        return df, applied
    return df[mask].reset_index(drop=True), applied


def describe_filter(spec):
    col, kind = spec['col'], spec['kind']
    if kind == 'number':
        lo, hi = spec['min'], spec['max']
        if lo is not None and hi is not None:
            return f"{col}: {_plain_num(lo)} to {_plain_num(hi)}"
        return f"{col} ≥ {_plain_num(lo)}" if lo is not None else f"{col} ≤ {_plain_num(hi)}"
    if kind == 'date':
        lo, hi = spec['min'], spec['max']
        if lo and hi:
            return f"{col}: {lo} to {hi}"
        return f"{col} from {lo}" if lo else f"{col} until {hi}"
    if 'values' in spec:
        shown = ', '.join(spec['values'][:3])
        extra = len(spec['values']) - 3
        return f"{col}: {shown}" + (f" +{extra} more" if extra > 0 else '')
    return f"{col} contains “{spec['text']}”"


def _empty_chart_entry(idx, x_col, y_col, config):
    """Placeholder card used when the filters leave no rows to chart."""
    return {
        'index': idx, 'x_col': x_col, 'y_col': y_col,
        'chart_type': config.get('chart_type', 'auto'),
        'period': config.get('period', 'none'), 'agg': config.get('agg', 'sum'),
        'calc': config.get('calc'), 'calc_label': '',
        'x_calc': config.get('x_calc'), 'x_calc_label': '',
        'sort': config.get('sort', 'default'),
        'topn': config.get('topn', 'all'),
        'topn_n': config.get('topn_n', 10),
        'display_title': f"{x_col} vs {y_col}",
        'active_chart_type': config.get('chart_type', 'auto'),
        'auto_detected_label': None,
        'plot_html': '<p style="margin:30px 0;color:var(--text-muted);">'
                     'No rows match the current filters. Adjust or clear them to see this chart.</p>',
        'insights_text': '',
        'insight_is_custom': False,
    }


# ---------------------------------------------------------------------
# SORT / RANK / TOP-BOTTOM N  (for charts that show one value per group)
# ---------------------------------------------------------------------

def _rank_int(value, default=10):
    try:
        n = int(float(value))
    except (TypeError, ValueError):
        n = default
    return max(1, min(MAX_TOP_N, n))


def _read_rank_post(post):
    sort = post.get('sort', 'default')
    topn = post.get('topn', 'all')
    return (sort if sort in SORT_OPTIONS else 'default',
            topn if topn in TOPN_MODES else 'all',
            _rank_int(post.get('topn_n')))


def _rank_signature(sort, topn, n):
    """N only matters when Top/Bottom is switched on."""
    return (sort, topn, n if topn != 'all' else None)


def _sort_by_label(frame, col, ascending):
    if frame[col].dtype.kind in 'bifc':
        return frame.sort_values(col, ascending=ascending, kind='mergesort')
    return frame.sort_values(col, ascending=ascending, kind='mergesort',
                             key=lambda s: s.astype(str).str.casefold())


def prepare_ranking(df, x_col, y_col, config, plot):
    """Apply the chart's 'Sort' and 'Top / Bottom N' settings.

    Only touches charts that show one number per group (bar, horizontal bar,
    pie, including calculated values). Anything else comes back unchanged.
    """
    sort = config.get('sort', 'default')
    mode = config.get('topn', 'all')
    sort = sort if sort in SORT_OPTIONS else 'default'
    mode = mode if mode in TOPN_MODES else 'all'
    if sort == 'default' and mode == 'all':
        return plot
    if plot.x_spec is not None:
        return plot                      # per-row X values (scatter / histogram / box)

    if plot.y_spec is not None:
        # A calculated value is already one result per group
        if plot.draw_type not in RANKABLE_TYPES:
            return plot
        grouped, x_name, y_name = plot.df.copy(), plot.x, plot.y
        draw_type, measure = plot.draw_type, 'value'
    elif plot.df is not df:
        return plot                      # date grouping: the order is time, leave it
    else:
        chart_type = config.get('chart_type', 'auto')
        if x_col not in df.columns or y_col not in df.columns:
            return plot
        draw_type = detect_chart_type(df, x_col, y_col) if chart_type == 'auto' else chart_type
        if draw_type not in RANKABLE_TYPES:
            return plot
        x_name = x_col
        if x_col != y_col and df[y_col].dtype.kind in 'bifc':
            y_name, measure = y_col, 'total'
            grouped = df.groupby(x_col)[y_col].sum().reset_index()
        else:
            y_name, measure = COUNT_LABEL, 'count'
            grouped = df[x_col].value_counts().rename(y_name).rename_axis(x_col).reset_index()

    total_groups = len(grouped)
    if total_groups == 0:
        return plot
    if draw_type == 'pie':
        sort = 'default'                 # a pie is always drawn largest slice first

    n = _rank_int(config.get('topn_n'))
    notes, limited = [], False
    if mode in ('top', 'bottom'):
        grouped = grouped.sort_values(y_name, ascending=(mode == 'bottom'), kind='mergesort')
        if n < total_groups:
            grouped = grouped.head(n)
            limited = True
            notes.append(f"{'Top' if mode == 'top' else 'Bottom'} {n} of {total_groups} by {measure}")

    if sort == 'value_desc':
        grouped = grouped.sort_values(y_name, ascending=False, kind='mergesort')
    elif sort == 'value_asc':
        grouped = grouped.sort_values(y_name, ascending=True, kind='mergesort')
    elif sort == 'label_asc':
        grouped = _sort_by_label(grouped, x_name, True)
    elif sort == 'label_desc':
        grouped = _sort_by_label(grouped, x_name, False)
    if sort != 'default':
        notes.append(f"sorted {SORT_NOTES[sort]}")

    note = ', '.join(notes)
    if plot.note and note:
        note = f"{plot.note}; {note}"
    elif plot.note:
        note = plot.note
    return PlotData(grouped.reset_index(drop=True), x_name, y_name, note, draw_type,
                    None, plot.y_spec, 'limited' if limited else 'sorted')


def _ranked_insight_note(plot):
    """Top/Bottom N hides groups, so say that 'highest', 'lowest' and shares refer to the groups shown."""
    if plot.ranked == 'limited':
        return (f" Note: this chart shows only {plot.note}, so the highest/lowest "
                f"groups and shares above refer to the groups shown.")
    return ''


# ---------------------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------------------

def _kpi_allowed_aggs(kind):
    if kind == 'number':
        return tuple(KPI_AGGS)
    if kind == 'date':
        return ('min', 'max', 'count', 'nunique', 'mode')
    return ('count', 'nunique', 'mode')


def normalize_kpi(df, raw):
    if not isinstance(raw, dict):
        return None
    col, agg = raw.get('col'), raw.get('agg')
    if col == RECORDS_COL:
        return {'col': RECORDS_COL, 'agg': 'count'}
    if col not in df.columns:
        return None
    if agg not in _kpi_allowed_aggs(column_kind(df[col])):
        return None
    return {'col': col, 'agg': agg}


def default_kpi_specs(df):
    """A sensible starting set for any file: record count, the totals of the
    first few real measures, and how many distinct values a category has."""
    specs = [{'col': RECORDS_COL, 'agg': 'count'}]
    measures = [c for c in df.select_dtypes(include='number').columns
                if not _is_identifier_like(df[c], c)]
    for col in measures[:3]:
        specs.append({'col': col, 'agg': 'sum'})
    for col in df.columns:
        series = df[col]
        if column_kind(series) != 'text':
            continue
        n = series.nunique(dropna=True)
        if 2 <= n <= max(2, len(series) * 0.5):
            specs.append({'col': col, 'agg': 'nunique'})
            break
    return specs[:5]


def _kpi_title(spec, kind):
    col, agg = spec['col'], spec['agg']
    if col == RECORDS_COL:
        return 'Records'
    name = _pretty(col)
    if agg == 'sum':
        return f"Total {name}"
    if agg == 'mean':
        return f"Average {name}"
    if agg == 'median':
        return f"Median {name}"
    if agg == 'min':
        return f"{'Earliest' if kind == 'date' else 'Minimum'} {name}"
    if agg == 'max':
        return f"{'Latest' if kind == 'date' else 'Maximum'} {name}"
    if agg == 'count':
        return f"{name} (non-blank)"
    if agg == 'nunique':
        return f"Distinct {name}"
    return f"Most common {name}"


def _kpi_value(df, spec):
    col, agg = spec['col'], spec['agg']
    if col == RECORDS_COL:
        return len(df)
    if len(df) == 0:
        return None
    series = df[col]
    if agg == 'count':
        return int(series.count())
    if agg == 'nunique':
        return int(series.nunique(dropna=True))
    if agg == 'mode':
        mode = series.mode(dropna=True)
        return str(mode.iloc[0]) if len(mode) else None
    if column_kind(series) == 'date':
        dates = _to_datetime(series).dropna()
        if dates.empty:
            return None
        return (dates.min() if agg == 'min' else dates.max()).strftime('%Y-%m-%d')
    result = getattr(pd.to_numeric(series, errors='coerce'), agg)()
    return None if pd.isna(result) else float(result)


def _fmt_kpi(value, symbol='', count_like=False, compact=True):
    if value is None:
        return '—'
    if isinstance(value, str):
        return value
    sign = '-' if value < 0 else ''
    v = abs(value)
    if count_like:
        body = f"{v:,.0f}"
    elif compact and v >= 1e9:
        body = f"{v / 1e9:,.2f}B"
    elif compact and v >= 1e6:
        body = f"{v / 1e6:,.2f}M"
    elif float(v).is_integer():
        body = f"{v:,.0f}"
    else:
        body = f"{v:,.2f}"
    return f"{sign}{symbol}{body}"


def compute_kpis(df, baseline, specs, currency, filters_active):
    """Build the KPI cards. When filters are on, each card also shows the
    figure for the whole file so you can see what the filter changed."""
    cards = []
    for i, raw in enumerate(specs or []):
        spec = normalize_kpi(baseline, raw)
        if spec is None:
            continue
        col, agg = spec['col'], spec['agg']
        kind = 'number' if col == RECORDS_COL else column_kind(baseline[col])
        count_like = col == RECORDS_COL or agg in ('count', 'nunique')
        symbol = '' if (count_like or agg == 'mode') else (currency or {}).get(col, '')

        value = _kpi_value(df, spec)
        card = {
            'index': i,
            'label': _kpi_title(spec, kind),
            'value': _fmt_kpi(value, symbol, count_like),
            'full': _fmt_kpi(value, symbol, count_like, compact=False),
            'sub': '', 'delta': '', 'delta_class': '',
        }

        if filters_active:
            base = _kpi_value(baseline, spec)
            card['sub'] = f"All data: {_fmt_kpi(base, symbol, count_like)}"
            if isinstance(value, (int, float)) and isinstance(base, (int, float)) and base:
                if agg in ('sum', 'count', 'nunique'):
                    card['delta'] = f"{value / base * 100:.1f}% of all"
                else:
                    change = (value - base) / abs(base) * 100
                    card['delta'] = f"{change:+.1f}% vs all"
                    card['delta_class'] = 'up' if change > 0 else 'down' if change < 0 else ''
        cards.append(card)
    return cards


# ---------------------------------------------------------------------
# PIVOT TABLE  (rows x optional columns, any aggregation, rank + top/bottom N)
# ---------------------------------------------------------------------

def default_pivot_config(df):
    cols = list(df.columns)
    rows = ''
    for col in cols:
        series = df[col]
        if column_kind(series) == 'text' and 2 <= series.nunique(dropna=True) <= 50:
            rows = col
            break
    if not rows:
        rows = cols[0]
    measures = [c for c in df.select_dtypes(include='number').columns
                if c != rows and not _is_identifier_like(df[c], c)]
    return {
        'rows': rows, 'cols': '',
        'values': measures[0] if measures else RECORDS_COL,
        'agg': 'sum' if measures else 'count',
        'period': 'none', 'sort': 'total', 'direction': 'desc',
        'topn': 'all', 'n': 10,
    }


def normalize_pivot_config(df, raw):
    """Fill in defaults and drop anything that doesn't fit this dataset."""
    cfg = default_pivot_config(df)
    raw = raw if isinstance(raw, dict) else {}
    columns = list(df.columns)

    if raw.get('rows') in columns:
        cfg['rows'] = raw['rows']
    if raw.get('cols') in columns and raw['cols'] != cfg['rows']:
        cfg['cols'] = raw['cols']
    if raw.get('values') == RECORDS_COL or raw.get('values') in columns:
        cfg['values'] = raw['values']
        cfg['agg'] = raw.get('agg') if raw.get('agg') in PIVOT_AGGS else 'sum'
    elif raw.get('agg') in PIVOT_AGGS and cfg['values'] != RECORDS_COL:
        cfg['agg'] = raw['agg']

    if cfg['values'] == RECORDS_COL:
        cfg['agg'] = 'count'
    elif column_kind(df[cfg['values']]) != 'number' and cfg['agg'] not in ('count', 'nunique'):
        cfg['agg'] = 'count'

    if raw.get('period') in TIME_PERIODS:
        cfg['period'] = raw['period']
    if raw.get('sort') in PIVOT_SORTS:
        cfg['sort'] = raw['sort']
    if raw.get('direction') in ('asc', 'desc'):
        cfg['direction'] = raw['direction']
    if raw.get('topn') in TOPN_MODES:
        cfg['topn'] = raw['topn']
    cfg['n'] = _rank_int(raw.get('n', cfg['n']))
    return cfg


def _read_pivot_post(post):
    return {
        'rows': post.get('pv_rows', ''), 'cols': post.get('pv_cols', ''),
        'values': post.get('pv_values', ''), 'agg': post.get('pv_agg', ''),
        'period': post.get('pv_period', 'none'), 'sort': post.get('pv_sort', 'total'),
        'direction': post.get('pv_direction', 'desc'), 'topn': post.get('pv_topn', 'all'),
        'n': post.get('pv_n', 10),
    }


def _pivot_keys(df, col, period):
    """One label per row for a pivot axis; dates can be grouped into periods."""
    series = df[col]
    if period in TIME_PERIODS and period != 'none' and _looks_like_date(series):
        freq = TIME_PERIODS[period][1]
        periods = _to_datetime(series).dt.to_period(freq)
        labels = periods.map(lambda p: _period_label(p, freq) if pd.notna(p) else None)
        return labels.fillna('(blank)')
    if series.dtype.kind in 'ifc' and series.notna().all():
        return series
    return series.astype(str).where(series.notna(), '(blank)')


def _label_text(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _fmt_cell(value, symbol='', count_like=False):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return '—'
    sign = '-' if value < 0 else ''
    v = abs(value)
    body = f"{v:,.0f}" if (count_like or float(v).is_integer()) else f"{v:,.2f}"
    return f"{sign}{symbol}{body}"


def build_pivot(df, cfg, currency=None):
    """Build the pivot table. Returns (context_for_template, DataFrame_for_csv_download).

    Rank 1 is always the biggest row total, so a Bottom-N table shows its
    true positions (e.g. #46-#50) instead of 1-5.
    """
    rows_col, cols_col, values_col = cfg['rows'], cfg['cols'], cfg['values']
    agg = cfg['agg']
    period = cfg['period']

    if len(df) == 0:
        return {'error': 'No rows to build a pivot table from. Adjust or clear the filters.'}, None

    work = pd.DataFrame({'r': _pivot_keys(df, rows_col, period)})
    if cols_col:
        work['c'] = _pivot_keys(df, cols_col, period)

    if values_col == RECORDS_COL:
        work['v'] = 1
        agg_fn, value_label = 'sum', COUNT_LABEL
        count_like, symbol = True, ''
    else:
        if agg in ('count', 'nunique'):
            work['v'] = df[values_col]
        else:
            work['v'] = pd.to_numeric(df[values_col], errors='coerce')
        agg_fn = agg
        value_label = f"{PIVOT_AGGS[agg]} of {values_col}"
        count_like = agg in ('count', 'nunique')
        symbol = '' if count_like else (currency or {}).get(values_col, '')

    row_totals = work.groupby('r')['v'].agg(agg_fn)
    grand = work['v'].agg(agg_fn)
    additive = agg_fn in ('sum', 'count')

    cell_table, col_totals, col_order, notes = None, None, [], []
    if cols_col:
        cell_table = work.groupby(['r', 'c'])['v'].agg(agg_fn).unstack('c')
        col_totals = work.groupby('c')['v'].agg(agg_fn)
        col_order = list(cell_table.columns)
        if len(col_order) > MAX_PIVOT_COLS:
            keep = set(col_totals.sort_values(ascending=False).head(MAX_PIVOT_COLS).index)
            notes.append(f"Showing the {MAX_PIVOT_COLS} largest of {len(col_order)} '{cols_col}' "
                         f"columns; row totals still include all of them.")
            col_order = [c for c in col_order if c in keep]
        cell_table = cell_table.reindex(columns=col_order)
        if agg_fn in ('sum', 'count', 'nunique'):
            cell_table = cell_table.fillna(0)

    # ---- rank, Top/Bottom N, sort ----
    frame = pd.DataFrame({'total': row_totals})
    frame['rank'] = frame['total'].rank(method='min', ascending=False)
    n_all = len(frame)
    mode, n = cfg['topn'], cfg['n']
    basis = 'row total' if cols_col else 'value'

    if mode in ('top', 'bottom') and n < n_all:
        frame = frame.sort_values('total', ascending=(mode == 'bottom'), kind='mergesort').head(n)
        notes.insert(0, f"Showing the {'top' if mode == 'top' else 'bottom'} {len(frame)} of "
                        f"{n_all:,} '{rows_col}' rows by {basis}.")
    ascending = cfg['direction'] == 'asc'
    if cfg['sort'] == 'label':
        if frame.index.dtype.kind in 'bifc':
            frame = frame.sort_index(ascending=ascending, kind='mergesort')
        else:
            frame = frame.sort_index(ascending=ascending, kind='mergesort',
                                     key=lambda idx: idx.astype(str).str.casefold())
    else:
        frame = frame.sort_values('total', ascending=ascending, kind='mergesort')
    if mode == 'all' and len(frame) > MAX_PIVOT_ROWS:
        frame = frame.head(MAX_PIVOT_ROWS)
        notes.insert(0, f"Showing the first {MAX_PIVOT_ROWS} of {n_all:,} '{rows_col}' rows. "
                        f"Use Top / Bottom N or a filter to narrow it down.")
    subset = len(frame) < n_all

    totals_valid = row_totals.dropna()
    has_share = bool(additive and pd.notna(grand) and grand > 0 and (totals_valid >= 0).all())
    finite_totals = frame['total'].dropna()
    top_total = float(finite_totals.max()) if len(finite_totals) else 0.0

    shade_lo = shade_hi = None
    if cell_table is not None and cell_table.size:
        flat = cell_table.to_numpy(dtype=float).ravel()
        flat = flat[~np.isnan(flat)]
        if flat.size:
            shade_lo, shade_hi = float(flat.min()), float(flat.max())

    def shade(v):
        if shade_lo is None or v is None or np.isnan(v) or shade_hi == shade_lo:
            return ''
        alpha = 0.06 + 0.34 * (float(v) - shade_lo) / (shade_hi - shade_lo)
        return f"background: rgba(56,189,248,{alpha:.2f});"

    table_rows, export_rows = [], []
    for label, rec in frame.iterrows():
        total = rec['total']
        total_ok = pd.notna(total)
        cells, raw_cells = [], []
        for c in col_order:
            v = cell_table.loc[label, c]
            v = None if pd.isna(v) else float(v)
            cells.append({'text': _fmt_cell(v, symbol, count_like), 'style': shade(v)})
            raw_cells.append(v)

        bar = ''
        if not cols_col and total_ok and top_total > 0 and total >= 0:
            width = float(total) / top_total * 100
            bar = f"background: linear-gradient(90deg, rgba(56,189,248,0.30) {width:.1f}%, transparent {width:.1f}%);"
        share = (f"{float(total) / float(grand) * 100:.1f}%" if has_share and total_ok else '')

        rank = int(rec['rank']) if pd.notna(rec['rank']) else None
        table_rows.append({
            'rank': rank if rank is not None else '—',
            'label': _label_text(label),
            'cells': cells,
            'total': _fmt_cell(float(total) if total_ok else None, symbol, count_like),
            'style': bar,
            'share': share,
        })
        export_rows.append(
            [rank, _label_text(label)] + raw_cells
            + [float(total) if total_ok else None]
            + ([float(total) / float(grand) * 100 if total_ok else None] if has_share else [])
        )

    total_cells = []
    if cols_col:
        for c in col_order:
            v = col_totals.get(c)
            total_cells.append({'text': _fmt_cell(None if pd.isna(v) else float(v), symbol, count_like), 'style': ''})

    # Raw numbers behind the table, used to write the pivot's key insight and recommendation
    valid_totals = row_totals.dropna().sort_values(ascending=False, kind='mergesort')
    stats = {}
    if len(valid_totals):
        top_cell = None
        if cell_table is not None and cell_table.size:
            arr = cell_table.to_numpy(dtype=float)
            if np.isfinite(arr).any():
                ri, ci = np.unravel_index(np.nanargmax(arr), arr.shape)
                top_cell = (_label_text(cell_table.index[ri]), _label_text(cell_table.columns[ci]),
                            float(arr[ri, ci]))
        stats = {
            'n_rows': len(valid_totals), 'n_all': n_all, 'shown': len(frame), 'subset': subset,
            'top': (_label_text(valid_totals.index[0]), float(valid_totals.iloc[0])),
            'bottom': (_label_text(valid_totals.index[-1]), float(valid_totals.iloc[-1])),
            'top3_total': float(valid_totals.head(3).sum()),
            'mean': float(valid_totals.mean()),
            'grand': float(grand) if pd.notna(grand) else None,
            'has_share': has_share, 'top_cell': top_cell,
            'symbol': symbol, 'count_like': count_like,
        }

    ctx = {
        'row_label': rows_col,
        'col_label': cols_col,
        'value_label': value_label,
        'columns': [_label_text(c) for c in col_order],
        'rows': table_rows,
        'has_share': has_share,
        'totals': {
            'label': ('Grand total' if additive else 'Overall') + (f" (all {n_all:,} rows)" if subset else ''),
            'cells': total_cells,
            'total': _fmt_cell(None if pd.isna(grand) else float(grand), symbol, count_like),
            'share': '100.0%' if has_share else '',
        },
        'notes': notes,
        'rank_note': f"Rank 1 = highest {basis}; tied rows share a rank.",
        'stats': stats,
    }

    header = ['Rank', rows_col] + [_label_text(c) for c in col_order] + ['Total' if cols_col else value_label]
    if has_share:
        header.append('% of total')
    return ctx, pd.DataFrame(export_rows, columns=header)


def _pivot_fmt(stats):
    symbol = stats.get('symbol', '')
    count_like = stats.get('count_like', False)
    return lambda v: _fmt_cell(v, symbol, count_like)


def generate_pivot_insight(pivot):
    """Plain-language key insight for the pivot table: who leads, who trails,
    how concentrated the total is. Works on any dataset (no column names assumed)."""
    try:
        s = (pivot or {}).get('stats')
        if not s or not s.get('n_rows'):
            return ''
        fmt = _pivot_fmt(s)
        row_label, col_label = pivot['row_label'], pivot.get('col_label')
        measure = pivot['value_label']
        n = s['n_rows']
        top_label, top_val = s['top']
        low_label, low_val = s['bottom']
        grand = s['grand']

        if n == 1:
            parts = [f"{measure} has only one '{row_label}' group, '{top_label}', at {fmt(top_val)}."]
        elif top_val == low_val:
            parts = [f"Every '{row_label}' group has the same {measure} ({fmt(top_val)}), so no group stands out."]
        elif s['has_share'] and grand:
            parts = [
                f"'{top_label}' ranks first in {measure} by {row_label} at {fmt(top_val)} "
                f"({top_val / grand * 100:.1f}% of the {fmt(grand)} total), while '{low_label}' "
                f"ranks last at {fmt(low_val)} ({low_val / grand * 100:.1f}%)."
            ]
            if n >= 5:
                parts.append(
                    f"The top 3 of {n:,} {row_label} groups account for "
                    f"{s['top3_total'] / grand * 100:.1f}% of the total."
                )
        else:
            parts = [
                f"'{top_label}' is highest in {measure} by {row_label} at {fmt(top_val)} and "
                f"'{low_label}' is lowest at {fmt(low_val)}, a gap of {fmt(top_val - low_val)}. "
                f"The average across the {n:,} {row_label} groups is {fmt(s['mean'])}."
            ]

        if col_label and s.get('top_cell'):
            r, c, v = s['top_cell']
            parts.append(f"The largest single cell is '{r}' under '{c}' ({col_label}) at {fmt(v)}.")
        if s['subset']:
            parts.append(
                f"The table shows only {s['shown']:,} of {s['n_all']:,} rows, "
                f"but the figures above cover all of them."
            )
        return ' '.join(parts)
    except Exception:
        return ''


def generate_pivot_recommendation(pivot):
    """Recommendation that fits what the pivot table actually shows."""
    try:
        s = (pivot or {}).get('stats')
        if not s or not s.get('n_rows'):
            return ''
        row_label, col_label = pivot['row_label'], pivot.get('col_label')
        measure = pivot['value_label']
        n = s['n_rows']
        top_label, top_val = s['top']
        low_label, low_val = s['bottom']
        grand = s['grand']

        if n == 1:
            return (f"Only one '{row_label}' group is available, so there is nothing to compare yet. "
                    f"Pick a row field with more distinct values or loosen the filters.")
        if top_val == low_val:
            return (f"All '{row_label}' groups have the same {measure}, so this layout gives no basis "
                    f"for prioritizing. Try a different row field or measure.")

        recs = []
        if s['has_share'] and grand:
            top_share = top_val / grand * 100
            low_share = low_val / grand * 100
            top3_share = s['top3_total'] / grand * 100
            if top_share >= 40:
                recs.append(
                    f"'{top_label}' alone makes up {top_share:.0f}% of the total, so results lean heavily "
                    f"on one {row_label}. Look into what drives it, treat that dependence as a risk, "
                    f"and plan ways to grow the other groups."
                )
            elif n >= 6 and top3_share >= 70:
                recs.append(
                    f"The top three {row_label} groups carry {top3_share:.0f}% of the total. Protect them, "
                    f"and check whether the remaining {n - 3:,} groups justify the effort they take."
                )
            else:
                recs.append(
                    f"The total is spread fairly evenly across {row_label} groups (largest share "
                    f"{top_share:.0f}%). Study what separates '{top_label}' from '{low_label}' and "
                    f"apply what works to the weaker groups."
                )
            if n >= 4 and low_share < 5:
                recs.append(
                    f"'{low_label}' contributes only {low_share:.1f}%, so decide whether it needs "
                    f"attention or should be merged into a broader group."
                )
        else:
            recs.append(
                f"'{top_label}' is highest and '{low_label}' lowest for {measure}. Find out what "
                f"separates them before acting, and check how many records sit behind each group "
                f"(try 'Number of records') so small groups do not mislead you."
            )

        if col_label and s.get('top_cell'):
            r, c, _ = s['top_cell']
            recs.append(
                f"Find out why '{r}' under '{c}' is the strongest combination and whether it can be "
                f"repeated, then review the lightest-shaded cells for weak spots."
            )
        if s['subset']:
            recs.append(
                f"Clear the Top/Bottom N setting before drawing conclusions about the whole dataset, "
                f"because the table shows only {s['shown']:,} of {s['n_all']:,} rows."
            )
        return ' '.join(recs)
    except Exception:
        return ''


# ---------------------------------------------------------------------
# PAGE LAYOUT: charts and pivot tables share one order the person can change
# ---------------------------------------------------------------------
# A block is addressed by a token: 'c<n>' for chart n, 'p<n>' for pivot n
# (positions in the saved lists). session['layout'] holds the tokens in the
# order they appear on the page and in the PDF.

def _to_int(value, default=-1):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _session_pivots(request):
    """The saved pivot tables: [{'cfg': settings or None, 'insight': text or None,
    'recommendation': text or None}]. No saved list means one default pivot, and
    sessions from before multiple pivots existed are carried over."""
    saved = request.session.get('pivots')
    if saved is None:
        saved = [{
            'cfg': request.session.get('pivot'),
            'insight': request.session.get('pivot_insight_override'),
            'recommendation': request.session.get('pivot_recommendation_override'),
        }]
    pivots = []
    for entry in saved:
        entry = entry if isinstance(entry, dict) else {}
        pivots.append({
            'cfg': entry.get('cfg') if isinstance(entry.get('cfg'), dict) else None,
            'insight': entry.get('insight') or None,
            'recommendation': entry.get('recommendation') or None,
        })
    return pivots


def _normalize_layout(layout, n_charts, n_pivots):
    """Every chart and pivot exactly once. Unknown or repeated tokens are dropped;
    anything missing goes at the end (pivot tables first, then charts)."""
    valid = {f'c{i}' for i in range(n_charts)} | {f'p{i}' for i in range(n_pivots)}
    ordered, seen = [], set()
    for token in layout or []:
        if token in valid and token not in seen:
            ordered.append(token)
            seen.add(token)
    for token in [f'p{i}' for i in range(n_pivots)] + [f'c{i}' for i in range(n_charts)]:
        if token not in seen:
            ordered.append(token)
    return ordered


def _layout_for(request):
    n_charts = len(request.session.get('charts') or [])
    return _normalize_layout(request.session.get('layout'), n_charts, len(_session_pivots(request)))


def _chart_order(layout):
    return [token for token in layout if token.startswith('c')]


def _block_anchor(token):
    if token and token[0] == 'c' and token[1:].isdigit():
        return f'chart-{token[1:]}'
    if token and token[0] == 'p' and token[1:].isdigit():
        return f'pivot-{token[1:]}'
    return ''


def _renumber_chart_refs(lines, old_order, new_order):
    """Saved recommendations start with 'Chart N - ...'. When charts are moved, keep each
    line attached to the same chart by rewriting N to the chart's new position."""
    if not lines:
        return lines
    mapping = {old_order.index(tok) + 1: new_order.index(tok) + 1 for tok in new_order if tok in old_order}
    if all(old == new for old, new in mapping.items()):
        return lines
    fixed = []
    for line in lines:
        match = re.match(r'^Chart (\d+)( - .*)$', str(line), re.DOTALL)
        if match:
            line = f"Chart {mapping.get(int(match.group(1)), int(match.group(1)))}{match.group(2)}"
        fixed.append(line)
    return fixed


def _order_charts(request, rendered_charts):
    """Charts in the order chosen on the page, each numbered 1..n in that order."""
    by_index = {chart['index']: chart for chart in rendered_charts}
    ordered = []
    for token in _layout_for(request):
        if token[0] == 'c' and by_index.get(int(token[1:])) is not None:
            chart = by_index[int(token[1:])]
            chart['number'] = len(ordered) + 1
            ordered.append(chart)
    return ordered


def _build_blocks(request, charts_in_order, pivots, skip_broken_pivots=False):
    """The page/PDF sequence: one entry per chart or pivot table, in layout order."""
    chart_by_index = {chart['index']: chart for chart in charts_in_order}
    pivot_by_index = {pivot['index']: pivot for pivot in pivots}
    blocks = []
    for token in _layout_for(request):
        index = int(token[1:])
        if token[0] == 'c':
            item = chart_by_index.get(index)
            if item is not None:
                blocks.append({'kind': 'chart', 'token': token, 'chart': item, 'pivot': None})
        else:
            item = pivot_by_index.get(index)
            if item is not None and not (skip_broken_pivots and item['table'].get('error')):
                blocks.append({'kind': 'pivot', 'token': token, 'chart': None, 'pivot': item})
    for position, block in enumerate(blocks, 1):
        block['position'] = position
        block['is_first'] = position == 1
        block['is_last'] = position == len(blocks)
    return blocks


def _pivot_title(table):
    if not table or table.get('error'):
        return 'Pivot table'
    title = f"{table.get('value_label', '')} by {table.get('row_label', '')}".strip()
    if table.get('col_label'):
        title += f" \u00d7 {table['col_label']}"
    return title


def _next_pivot_config(df, existing_cfgs):
    """Starting settings for a new pivot table: a different Rows field from the
    ones already on the page when the data has one, so it adds something new."""
    used = {cfg.get('rows') for cfg in existing_cfgs}
    base = default_pivot_config(df)
    for col in df.columns:
        series = df[col]
        if col not in used and column_kind(series) == 'text' and 2 <= series.nunique(dropna=True) <= 50:
            base['rows'] = col
            break
    else:
        for col in df.columns:
            if col not in used and column_kind(df[col]) == 'date':
                base['rows'], base['period'] = col, 'M'
                break
    return normalize_pivot_config(df, base)


def _download_pivot(request, file_path):
    """CSV download of the pivot table exactly as shown on the page."""
    try:
        df_full = read_dataset(file_path)
        df, _ = apply_filters(df_full, request.session.get('filters', []))
        pivots = _session_pivots(request)
        index = _to_int(request.GET.get('i'), 0)
        if not 0 <= index < len(pivots):
            return HttpResponse('That pivot table no longer exists.', status=404)
        cfg = normalize_pivot_config(df_full, pivots[index].get('cfg'))
        ctx, export_df = build_pivot(df, cfg, request.session.get('currency_symbols', {}))
    except Exception as e:
        return HttpResponse(f"Could not build the pivot table: {e}", status=500)
    if export_df is None:
        return HttpResponse(ctx.get('error', 'No pivot table available.'), status=400)
    number = _layout_for(request)
    number = [int(tok[1:]) for tok in number if tok[0] == 'p'].index(index) + 1
    response = HttpResponse('\ufeff' + export_df.to_csv(index=False), content_type='text/csv; charset=utf-8')
    response['Content-Disposition'] = f'attachment; filename="cascade_pivot_table_{number}.csv"'
    return response


def _dashboard_extras(request, df_full, df, applied, currency):
    """Everything the page needs for the filter bar, KPI cards and pivot table.
    A problem in one panel shows up in that panel instead of breaking the page."""
    extras = {
        'filter_chips': [{'index': i, 'label': describe_filter(s)} for i, s in applied],
        'filters_active': bool(applied),
        'rows_total': len(df_full),
        'rows_shown': len(df),
        'sort_options': list(SORT_OPTIONS.items()),
        'topn_modes': list(TOPN_MODES.items()),
        'kpi_aggs': list(KPI_AGGS.items()),
        'pivot_aggs': list(PIVOT_AGGS.items()),
        'pivot_sorts': list(PIVOT_SORTS.items()),
        'records_col': RECORDS_COL,
        'max_top_n': MAX_TOP_N,
    }
    try:
        extras['column_meta'] = build_filter_meta(df_full)
    except Exception:
        extras['column_meta'] = {}

    specs = request.session.get('kpis')
    extras['kpi_is_custom'] = specs is not None
    try:
        if specs is None:
            specs = default_kpi_specs(df_full)
        extras['kpis'] = compute_kpis(df, df_full, specs, currency, bool(applied))
    except Exception:
        extras['kpis'] = []

    # Pivot tables (any number), each with its own settings and editable insight + recommendation
    layout = _layout_for(request)
    pivot_order = [int(token[1:]) for token in layout if token[0] == 'p']
    pivots = []
    for index, entry in enumerate(_session_pivots(request)):
        try:
            cfg = normalize_pivot_config(df_full, entry.get('cfg'))
            table, _ = build_pivot(df, cfg, currency)
        except Exception as e:
            cfg = {}
            table = {'error': f"Could not build the pivot table: {e}"}
        ok = not table.get('error')
        insight_override = entry.get('insight')
        recommendation_override = entry.get('recommendation')
        pivots.append({
            'index': index,
            'number': pivot_order.index(index) + 1 if index in pivot_order else index + 1,
            'cfg': cfg,
            'table': table,
            'title': _pivot_title(table),
            'insight': insight_override or (generate_pivot_insight(table) if ok else ''),
            'recommendation': recommendation_override or (generate_pivot_recommendation(table) if ok else ''),
            'insight_is_custom': bool(insight_override),
            'recommendation_is_custom': bool(recommendation_override),
        })
    extras['pivots'] = pivots
    extras['pivot_custom_text'] = any(p['insight_is_custom'] or p['recommendation_is_custom'] for p in pivots)
    return extras


def _chart_display_title(x_col, y_col, plot, active_type):
    """'X vs Y' heading for the PDF, using the calculation's name when there is one."""
    x_name = plot.x if plot.x_spec else x_col
    y_name = plot.y if plot.y_spec else y_col
    if plot.x_spec and active_type in ('histogram', 'box'):
        return x_name
    return f"{x_name} vs {y_name}"


def _label_calc_axis(ax, drawn_type, label):
    """Name the value axis after the calculation (bars would say 'Total ...')."""
    if drawn_type == 'pie':
        return
    if drawn_type == 'horizontal_bar':
        ax.set_xlabel(label)
    else:
        ax.set_ylabel(label)


def _chart_source_cols(chart):
    """Dataset columns a chart depends on (used for the missing-data advice)."""
    cols = []
    # (setting, label it produces when applied, plain column fallback)
    for key, label_key, plain in (('x_calc', 'x_calc_label', 'x_col'),
                                  ('calc', 'calc_label', 'y_col')):
        spec = chart.get(key)
        if chart.get(label_key) and isinstance(spec, dict) and spec.get('on'):
            cols.append(spec.get('a'))
            if spec.get('b_kind') == 'col':
                cols.append(spec.get('b'))
        else:
            cols.append(chart.get(plain))
    return cols


def _calc_series(plot_df, x_col, y_col):
    data = plot_df[[x_col, y_col]].dropna()
    values = data[y_col].to_numpy(dtype=float)
    labels = [str(v) for v in data[x_col]]
    return values, labels


def generate_calc_insights(plot_df, x_col, y_col, active_type, spec):
    """Insight text for a calculated Y value (totals-based wording would be wrong here)."""
    values, labels = _calc_series(plot_df, x_col, y_col)
    unit = '%' if spec['pct'] else ''

    def fmt(v):
        return f"{v:,.2f}{unit}"

    if len(values) == 0:
        return f"No valid values could be calculated for {y_col} by {x_col}."

    how = calc_note(spec)
    if len(values) == 1:
        return f"{y_col} ({how}) is {fmt(values[0])} for {labels[0]}."

    hi, lo = int(values.argmax()), int(values.argmin())

    if active_type in ('line', 'area'):
        first, last = float(values[0]), float(values[-1])
        direction = "rose" if last > first else "fell" if last < first else "stayed flat"
        text = (
            f"{y_col} ({how}) {direction} from {fmt(first)} ({labels[0]}) to "
            f"{fmt(last)} ({labels[-1]}). The highest value was {fmt(values[hi])} "
            f"({labels[hi]}) and the lowest was {fmt(values[lo])} ({labels[lo]})."
        )
        if len(values) >= 6:
            half = len(values) // 2
            early = float(values[:half].mean())
            late = float(values[-half:].mean())
            if early != 0:
                change = (late - early) / abs(early) * 100
                text += (
                    f" Overall, the second half averages {abs(change):.1f}% "
                    f"{'higher' if change >= 0 else 'lower'} than the first half."
                )
        return text

    gap = float(values[hi] - values[lo])
    gap_text = f"{gap:,.2f} percentage points" if spec['pct'] else f"{gap:,.2f}"
    return (
        f"{y_col} ({how}) is highest for '{labels[hi]}' at {fmt(values[hi])} and "
        f"lowest for '{labels[lo]}' at {fmt(values[lo])}, a gap of {gap_text}. "
        f"The simple average across the {len(values)} {x_col} groups is "
        f"{fmt(values.mean())}."
    )


def _recommend_for_calc(plot_df, x_col, y_col, active_type, spec):
    """Recommendations for a chart that plots a calculated value."""
    values, labels = _calc_series(plot_df, x_col, y_col)
    if len(values) < 2:
        return []

    unit = '%' if spec['pct'] else ''
    hi, lo = int(values.argmax()), int(values.argmin())
    recs = [
        f"The widest difference in this chart is between '{labels[hi]}' "
        f"({values[hi]:,.2f}{unit}) and '{labels[lo]}' ({values[lo]:,.2f}{unit}). "
        f"Find out what drives that gap before acting on the average "
        f"({values.mean():,.2f}{unit})."
    ]
    recs.append(
        f"This chart uses the {calc_note(spec)}. Use the same method when "
        f"comparing with other reports so the figures stay comparable."
    )
    return recs


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

def draw_chart(ax, fig, df, x_col, y_col, chart_type, num_cols, keep_order=False, by_time=False):
    active_chart_type = detect_chart_type(df, x_col, y_col) if chart_type == 'auto' else chart_type

    if active_chart_type == 'bar':
        if df[y_col].dtype.kind in 'bifc':
            grouped = df.groupby(x_col, as_index=False, sort=not keep_order)[y_col].sum()
            if not keep_order:
                grouped = grouped.head(20)
            labels = grouped[x_col].astype(str).tolist()
            ax.bar(labels, grouped[y_col], color='#4f46e5', edgecolor='#312e81')
            ax.set_ylabel(y_col if (y_col == COUNT_LABEL or by_time) else f"Total {y_col}")
            if by_time and len(labels) > 24:
                # a long timeline: label about a dozen of the bars so they stay readable
                step = int(np.ceil(len(labels) / 12))
                ticks = list(range(0, len(labels), step))
                ax.set_xticks(ticks)
                ax.set_xticklabels([labels[i] for i in ticks])
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
            grouped = df.groupby(x_col, as_index=False, sort=not keep_order)[y_col].sum()
            if not keep_order:
                grouped = grouped.head(15)
            ax.barh(grouped[x_col].astype(str), grouped[y_col], color='#0284c7', edgecolor='#075985')
            ax.set_xlabel(y_col if y_col == COUNT_LABEL else f"Total {y_col}")
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

def generate_insights(df, x_col, y_col, chart_type, num_cols, currency=None, measure='sum'):
    """
    Generate insights that directly correspond to the data,
    parameters, aggregation, and chart type being displayed.

    `currency` maps column name -> currency symbol (detected in the cleaner)
    so money columns are shown as e.g. ₱1,250.00 instead of 1250.00.
    `measure` is how a date-grouped bar was combined (sum, mean, count, max,
    min) so the wording says "average" rather than "total" when it should.
    """
    currency = currency or {}

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
                f"{x_col} has a mean of {_money(mean, x_col, currency)} and a median of {_money(median, x_col, currency)}. "
                f"Values range from {_money(min_val, x_col, currency)} to {_money(max_val, x_col, currency)}, "
                f"with a standard deviation of {_money(std, x_col, currency)}. "
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
                f"{y_col} {trend} from {_money(first, y_col, currency)} to {_money(last, y_col, currency)}, "
                f"a change of {change_text}. "
                f"The highest value was {_money(values.max(), y_col, currency)} at "
                f"{x_col} = {peak_x}, while the lowest value was "
                f"{_money(values.min(), y_col, currency)} at {x_col} = {lowest_x}."
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
                f"observations of {x_col}. The total value is {_money(total, y_col, currency)}, "
                f"with an average of {_money(average, y_col, currency)}. "
                f"The highest value was {_money(maximum, y_col, currency)} at "
                f"{x_col} = {max_x}, while the lowest was "
                f"{_money(minimum, y_col, currency)} at {x_col} = {min_x}."
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

                if measure in ('mean', 'max', 'min'):
                    word = {'mean': 'average', 'max': 'maximum', 'min': 'minimum'}[measure]
                    return (
                        f"'{highest_category}' has the highest {word} {y_col} "
                        f"at {_money(highest_value, y_col, currency)}. "
                        f"'{lowest_category}' has the lowest {word} at "
                        f"{_money(lowest_value, y_col, currency)}."
                    )
                word = 'number of records' if measure == 'count' else 'total'
                return (
                    f"'{highest_category}' has the highest {word} {y_col} "
                    f"at {_money(highest_value, y_col, currency)}, representing "
                    f"{highest_share:.1f}% of the overall total. "
                    f"'{lowest_category}' has the lowest {word} at "
                    f"{_money(lowest_value, y_col, currency)}."
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
                f"{measured_column} has a median of {_money(median, measured_column, currency)}, "
                f"with the middle 50% of values ranging from "
                f"{_money(q1, measured_column, currency)} to {_money(q3, measured_column, currency)} "
                f"(IQR: {_money(iqr, measured_column, currency)}). "
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
CURRENCY_PATTERN = r"[$€£¥₱₹₩₽₫₺฿]|php|usd|eur|gbp|jpy|inr|cad|aud"


def _to_numeric_currency(series):
    """Convert text like '$1,250.00', '₱1,250', 'PHP 1,250' or '(1,250.00)'
    to numbers. Parentheses are treated as negatives (accounting style).
    Anything that still isn't a number becomes NaN."""
    text = series.astype(str).str.strip()

    # Accounting negatives: (1,250.00) -> -1250.00
    negative = text.str.match(r"^\(.*\)$")
    text = text.str.replace(r"^\((.*)\)$", r"\1", regex=True)

    # Remove currency symbols/codes, commas, and any kind of space
    text = text.str.replace(CURRENCY_PATTERN, "", regex=True, flags=re.IGNORECASE)
    text = text.str.replace(r"[,\s\u00a0]", "", regex=True)

    converted = pd.to_numeric(text, errors="coerce")
    return converted.where(~negative, -converted.abs())


CURRENCY_CODE_SYMBOLS = {
    "php": "₱", "usd": "$", "eur": "€", "gbp": "£", "jpy": "¥",
    "inr": "₹", "cad": "CA$", "aud": "A$",
}


def _detect_currency_symbol(series, min_share=0.5):
    """Return the currency symbol most of a text column's values carry
    (e.g. '₱' for '₱1,250.00'), or None if fewer than `min_share` of the
    non-empty values have one. Currency codes like 'PHP' map to symbols."""
    text = series.dropna().astype(str)
    if text.empty:
        return None

    found = text.str.findall(CURRENCY_PATTERN, flags=re.IGNORECASE)
    symbols = []
    for matches in found:
        if matches:
            m = matches[0]
            symbols.append(m if len(m) == 1 else CURRENCY_CODE_SYMBOLS.get(m.lower(), m.upper()))

    if not symbols or len(symbols) / len(text) < min_share:
        return None
    return pd.Series(symbols).value_counts().idxmax()


def _money(value, column, currency, default_fmt=".2f"):
    """Format a number for display in insights. If `column` carries a
    currency symbol (detected during cleaning) show it with thousands
    separators, e.g. ₱1,250.00 or -₱1,250.00. Otherwise format the number
    with `default_fmt` exactly as before."""
    symbol = (currency or {}).get(column)
    if not symbol:
        return format(value, default_fmt)
    if value < 0:
        return f"-{symbol}{abs(value):,.2f}"
    return f"{symbol}{value:,.2f}"


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
    """Rewrite every column that really holds dates into one clean, sortable
    format, without changing what the dates mean:

      full dates         -> YYYY-MM-DD       (YYYY-MM-DD HH:MM:SS if a time of day is present)
      month-level dates  -> YYYY-MM          ("Jan 2024", "Jan-24", "01/2024"; no day is invented)

    Values that are not dates (bare month names, years, quarters, codes) never
    qualify, so those columns stay as they are. A value inside a date column
    that cannot be read as a date is kept as it was rather than blanked.

    Returns (df, details), one entry per rewritten column."""
    details = []

    for column in df.columns:
        series = df[column]
        is_datetime = pd.api.types.is_datetime64_any_dtype(series)
        if not (pd.api.types.is_string_dtype(series) or is_datetime):
            continue
        if not _looks_like_date(series):
            continue
        non_null = series.notna().sum()
        if non_null == 0:
            continue

        parsed = _to_datetime(series)
        readable = parsed.notna()
        if readable.sum() / non_null < 0.8:
            continue

        # Which shapes were written, and could they be day/month or month/day?
        month_form_only, ambiguous = False, False
        if not is_datetime:
            texts = pd.unique(series.dropna().astype(str).str.strip())
            dayfirst, ambiguous = _infer_dayfirst(texts)
            forms = {_parse_date_text(text, dayfirst)[1] for text in texts}
            month_form_only = forms <= {'month', None} and 'month' in forms

        stamps = parsed[readable]
        has_time = bool(((stamps.dt.hour != 0) | (stamps.dt.minute != 0) | (stamps.dt.second != 0)).any())
        first_of_month = bool(
            not has_time and not month_form_only
            and (stamps.dt.day == 1).all() and stamps.nunique() >= 3
        )
        if month_form_only or first_of_month:
            precision, fmt = "month", "%Y-%m"
        elif has_time:
            precision, fmt = "time", "%Y-%m-%d %H:%M:%S"
        else:
            precision, fmt = "day", "%Y-%m-%d"

        rewritten = parsed.dt.strftime(fmt)
        left_alone = (~readable) & series.notna()
        if left_alone.any():
            rewritten = rewritten.where(~left_alone, series.astype(str))
        df[column] = rewritten

        details.append({
            "column": column,
            "precision": precision,
            "first_of_month": first_of_month,
            "ambiguous": bool(ambiguous),
            "left_alone": int(left_alone.sum()),
        })

    return df, details


_DATE_NAME_WORDS = {'date', 'dates', 'datetime', 'timestamp'}
_EXCEL_DAY_RANGE = (25569, 73050)                  # 1970-01-01 .. 2099-12-31 as Excel day numbers
_UNIX_SECOND_RANGE = (315532800, 4102444800)       # 1980 .. 2100 as Unix seconds


def _name_says_date(name):
    """True when a column is called Date, Order Date, order_date, OrderDate,
    Timestamp... (whole words only, so "Update" or "Candidate" do not count)."""
    spaced = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(name))
    return bool(set(re.findall(r"[a-z]+", spaced.lower())) & _DATE_NAME_WORDS)


def _numbers_as_dates(series):
    """A numeric column that really holds dates written as plain numbers:
    Excel day numbers (45292 = 2024-01-01) or Unix timestamps. Returns
    (timestamps, 'excel' | 'unix') or None. Callers only ask for columns
    whose name says they are dates, and every value must fall in a plausible
    date range, so IDs, prices and years are never converted."""
    if series.dtype.kind not in 'iuf':
        return None
    values = pd.to_numeric(series, errors='coerce').astype('float64')
    present = values.dropna()
    if present.empty or not np.isfinite(present).all():
        return None
    low, high = float(present.min()), float(present.max())

    if _EXCEL_DAY_RANGE[0] <= low and high <= _EXCEL_DAY_RANGE[1]:
        stamps = pd.Timestamp('1899-12-30') + pd.to_timedelta(values, unit='D')
        return stamps.dt.round('s'), 'excel'
    if _UNIX_SECOND_RANGE[0] <= low and high <= _UNIX_SECOND_RANGE[1]:
        return pd.to_datetime(values, unit='s').dt.round('s'), 'unix'
    if _UNIX_SECOND_RANGE[0] * 1000 <= low and high <= _UNIX_SECOND_RANGE[1] * 1000:
        return pd.to_datetime(values, unit='ms').dt.round('s'), 'unix'
    return None


def _convert_numeric_dates(df):
    """Turn number-dates into real dates (the later date step then writes them
    in the standard format). Returns (df, report_lines)."""
    lines = []
    for column in df.columns:
        if not _name_says_date(column):
            continue
        found = _numbers_as_dates(df[column])
        if found is None:
            continue
        stamps, kind = found
        before = df[column].dropna().iloc[0]
        before_text = str(int(before)) if float(before).is_integer() else str(before)
        after = stamps.dropna().iloc[0]
        df[column] = stamps
        what = "Excel dates" if kind == "excel" else "Unix timestamps"
        lines.append(
            f"Read the numbers in '{column}' as {what}, so they are now real dates "
            f"(for example {before_text} became {after:%Y-%m-%d})."
        )
    return df, lines


def _describe_date_changes(details):
    """Plain-language report lines for what _normalize_dates did."""
    lines = []
    by_precision = {"day": [], "time": [], "month": []}
    for item in details:
        by_precision[item["precision"]].append(item["column"])

    if by_precision["day"]:
        lines.append(f"Standardized date formats to YYYY-MM-DD in {len(by_precision['day'])} column(s).")
    if by_precision["time"]:
        names = ", ".join(f"'{c}'" for c in by_precision["time"])
        lines.append(f"Standardized dates with a time of day to YYYY-MM-DD HH:MM:SS (the time is kept) in: {names}.")
    if by_precision["month"]:
        names = ", ".join(f"'{c}'" for c in by_precision["month"])
        lines.append(f"Kept month-level dates at month level (YYYY-MM, no day added) in: {names}.")

    for item in details:
        col = item["column"]
        if item["ambiguous"]:
            lines.append(
                f"Dates in '{col}' such as 03/04/2024 could be read either way; "
                f"they were read as month/day/year because nothing in the column shows otherwise."
            )
        if item["left_alone"]:
            lines.append(
                f"{_plural(item['left_alone'], 'value')} in '{col}' did not look like a date "
                f"and {'was' if item['left_alone'] == 1 else 'were'} left as written."
            )
    return lines


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


# ---------------------------------------------------------------------------
# DATA CLEANING
#
# Cleaning runs in stages so the user stays in control of anything that
# removes or changes real data:
#   1. "Clean Data" -> the automatic, deterministic steps run (empty rows,
#      duplicates, spacing, number/date formats, ...). No data is lost to
#      missing values at this point.
#   2. If any rows have missing values, the user is shown how many (counts
#      and percentages) and must choose: remove those rows, or review them
#      one by one and either type the missing values in by hand or remove
#      the row. Nothing is ever filled in automatically.
#   3. A confirmation step, enforced on the server, before anything is applied.
#   4. The decisions are applied to the dataset, the cleaned CSV is written to
#      disk (this is the file Visualize and the PDF read), and a record of the
#      choices is kept in the session so they can be shown in both places.
#
# The flow is stateless between requests: the automatic steps are
# deterministic, so the review page and the apply step each re-run them from
# the uploaded file and refer to rows by their original position.
# ---------------------------------------------------------------------------

# Strings that count as "no value" (standardized to missing during cleaning).
MISSING_MARKERS = ["", " ", "NA", "N/A", "na", "n/a", "null", "NULL", "-"]

# Strings pandas reads back as missing. A value typed in by the user must not
# be one of these, or it would turn back into a missing value when the saved
# CSV is read again by the visualizer.
NA_LIKE_STRINGS = {
    "", "-", "#n/a", "#n/a n/a", "#na", "-1.#ind", "-1.#qnan", "-nan",
    "1.#ind", "1.#qnan", "<na>", "n/a", "na", "null", "nan", "none",
}

CLEANING_STATE_KEYS = (
    "charts", "insights_overrides", "filters", "kpis", "pivot",
    "pivot_insight_override", "pivot_recommendation_override",
    "pivots", "layout", "summary_override", "recommendations_override",
)

KIND_LABELS = {
    "integer": "whole number",
    "number": "number",
    "date": "date, YYYY-MM-DD",
    "month": "month, YYYY-MM",
    "datetime": "date and time, YYYY-MM-DD HH:MM:SS",
    "text": "text",
}
# <input type> used for each kind in the manual review
KIND_INPUT_TYPES = {
    "integer": "number", "number": "number", "date": "date",
    "month": "month", "datetime": "datetime-local", "text": "text",
}


def _plural(count, word):
    return f"{count:,} {word}" if count == 1 else f"{count:,} {word}s"


def _fmt_pct(part, total):
    if not total:
        return "0%"
    value = 100.0 * part / total
    if part and value < 0.1:
        return "<0.1%"
    return f"{value:.1f}%"


def _prepare_dataset(file_path):
    """The automatic part of cleaning. Needs no input from the user and gives
    the same result every time for the same file, so it can be re-run safely.

    Returns (df, report, info). `df` keeps the original row positions as its
    index (0 = first data row), which is how rows are referred to later."""
    df = read_dataset(file_path).reset_index(drop=True)

    original_rows = len(df)
    original_columns = len(df.columns)
    report = []

    # 1. Remove completely empty rows
    empty_rows = int(df.isna().all(axis=1).sum())

    if empty_rows > 0:
        df = df.dropna(how="all")
        report.append(f"Removed {empty_rows} completely empty row(s).")

    # 2. Remove duplicate rows
    duplicate_count = int(df.duplicated().sum())

    if duplicate_count > 0:
        df = df.drop_duplicates()
        report.append(f"Removed {duplicate_count} duplicate row(s).")

    # 3. Normalize whitespace (trim + collapse repeated spaces)
    df, spaces_removed = _normalize_whitespace(df)

    if spaces_removed > 0:
        report.append(
            f"Normalized spacing in {spaces_removed} value(s) (trimmed and collapsed extra spaces)."
        )

    # 4. Standardize missing values
    missing_count = 0

    for column in df.columns:
        for value in MISSING_MARKERS:
            missing_count += int((df[column] == value).sum())

        df[column] = df[column].replace(MISSING_MARKERS, pd.NA)

    if missing_count > 0:
        report.append(f"Standardized {missing_count} missing/invalid value(s).")

    # 5. Convert numeric-looking columns (handles currency symbols too)
    numeric_columns = 0
    detected_currency = {}

    for column in df.columns:

        if pd.api.types.is_string_dtype(df[column]):

            converted = _to_numeric_currency(df[column])

            non_empty = df[column].notna().sum()

            if non_empty > 0:
                numeric_count = converted.notna().sum()

                if numeric_count / non_empty >= 0.8:
                    # Remember the currency symbol (if any) before it is stripped
                    symbol = _detect_currency_symbol(df[column])
                    if symbol:
                        detected_currency[column] = symbol

                    df[column] = converted
                    numeric_columns += 1

    if numeric_columns > 0:
        report.append(
            f"Converted {numeric_columns} column(s) to numeric format "
            f"(currency symbols and commas removed)."
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
        report.append(f"Standardized {changed_columns} column name(s).")

    # Carry detected currency symbols over to the (renamed) columns so the
    # visualizer can show them in insights, e.g. ₱1,250.00 instead of 1250.00
    rename_map = dict(zip(old_columns, df.columns))
    currency_symbols = {
        rename_map[col]: symbol
        for col, symbol in detected_currency.items()
        if col in rename_map
    }

    if currency_symbols:
        currency_desc = ", ".join(
            f"'{col}' ({symbol})" for col, symbol in currency_symbols.items()
        )
        report.append(
            f"Detected currency in: {currency_desc}. Values are stored as plain "
            f"numbers and the symbol is shown in the visualization insights."
        )

    # 7. Normalize categorical value casing (data-driven, no hardcoded columns)
    df, casing_changed, casing_columns = _normalize_categorical_casing(df)

    if casing_columns > 0:
        report.append(
            f"Standardized inconsistent capitalization in {casing_changed} "
            f"value(s) across {casing_columns} column(s)."
        )

    # 8. Dates saved as plain numbers (Excel day numbers) become real dates,
    #    then every date column is written in one standard format
    df, numeric_date_lines = _convert_numeric_dates(df)
    report.extend(numeric_date_lines)
    df, date_details = _normalize_dates(df)
    report.extend(_describe_date_changes(date_details))

    # 9. Restore whole-number formatting where a column is entirely integers
    df, tightened_columns = _tighten_numeric_types(df)

    if tightened_columns > 0:
        report.append(
            f"Restored whole-number formatting in {tightened_columns} numeric column(s)."
        )

    info = {
        "original_rows": original_rows,
        "original_columns": original_columns,
        "currency_symbols": currency_symbols,
        # Row n of the source file is data row `index + row_offset`
        # (spreadsheets/CSV have a header line; JSON records start at 1).
        "row_offset": 1 if os.path.splitext(file_path)[1].lower() in JSON_EXTENSIONS else 2,
    }
    return df, report, info


def _missing_overview(df):
    """How much is missing, overall and per column, plus the ids (original
    row positions) of every row with at least one missing value."""
    mask = df.isna()
    row_mask = mask.any(axis=1)
    total_rows, total_columns = df.shape
    rows_with_missing = int(row_mask.sum())
    cells_missing = int(mask.to_numpy().sum())

    columns = []
    for pos, name in enumerate(df.columns):
        count = int(mask.iloc[:, pos].sum())
        if count:
            columns.append({
                "name": name,
                "count": count,
                "pct": _fmt_pct(count, total_rows),
                "bar": max(2, round(100.0 * count / total_rows)) if total_rows else 0,
            })
    columns.sort(key=lambda c: -c["count"])

    return {
        "total_rows": total_rows,
        "total_columns": total_columns,
        "rows_with_missing": rows_with_missing,
        "rows_pct": _fmt_pct(rows_with_missing, total_rows),
        "complete_rows": total_rows - rows_with_missing,
        "cells_missing": cells_missing,
        "cells_pct": _fmt_pct(cells_missing, total_rows * total_columns),
        "columns": columns,
        "ids": [int(i) for i in df.index[row_mask.to_numpy()]],
    }


def _column_kinds(df):
    """What kind of value each column holds, so manual entries can be
    checked: 'integer', 'number', 'date' (YYYY-MM-DD), 'month' (YYYY-MM),
    'datetime' (YYYY-MM-DD HH:MM:SS) or 'text'."""
    kinds = []
    for pos in range(df.shape[1]):
        series = df.iloc[:, pos]
        if pd.api.types.is_bool_dtype(series):
            kinds.append("text")
        elif pd.api.types.is_integer_dtype(series):
            kinds.append("integer")
        elif pd.api.types.is_numeric_dtype(series):
            kinds.append("number")
        else:
            present = series.dropna().astype(str)
            kind = "text"
            if len(present):
                for candidate, shape in (
                    ("date", r"\d{4}-\d{2}-\d{2}"),
                    ("month", r"\d{4}-\d{2}"),
                    ("datetime", r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}"),
                ):
                    if present.str.fullmatch(shape).all():
                        kind = candidate
                        break
            kinds.append(kind)
    return kinds


def _validate_manual_value(raw, kind, series, canon_cache, cidx):
    """Check a value the user typed for a missing cell. Returns
    (clean_value, None) or (None, error_message)."""
    import datetime

    text = re.sub(r"\s+", " ", (raw or "").strip())
    if not text:
        return None, "Enter a value."
    if text.lower() in NA_LIKE_STRINGS:
        return None, "That would be read as a missing value. Enter a real value."

    if kind in ("integer", "number"):
        number = _to_numeric_currency(pd.Series([text])).iloc[0]
        if pd.isna(number) or not np.isfinite(number):
            return None, "Enter a number, for example 1250 or 1,250.50."
        if kind == "integer":
            if number % 1 != 0:
                return None, "This column holds whole numbers. Enter a whole number."
            return int(number), None
        return float(number), None

    if kind == "date":
        try:
            return datetime.datetime.strptime(text, "%Y-%m-%d").strftime("%Y-%m-%d"), None
        except ValueError:
            return None, "Enter a valid date as YYYY-MM-DD, for example 2024-03-15."

    if kind == "month":
        try:
            return datetime.datetime.strptime(text, "%Y-%m").strftime("%Y-%m"), None
        except ValueError:
            return None, "Enter a month as YYYY-MM, for example 2024-03."

    if kind == "datetime":
        for shape in ("%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M",
                      "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text, shape).strftime("%Y-%m-%d %H:%M:%S"), None
            except ValueError:
                continue
        return None, "Enter a date and time such as 2024-03-15 14:30."

    if len(text) > 200:
        return None, "Keep it under 200 characters."

    # Match the capitalization already used in the column ("south" -> "South")
    if cidx not in canon_cache:
        existing = series.dropna().astype(str).unique()
        canon_cache[cidx] = {v.lower(): v for v in existing}
    return canon_cache[cidx].get(text.lower(), text), None


def _collect_manual_decisions(df, kinds, ids, posted):
    """Read and validate the per-row decisions from the review form.
    Every row needs an explicit Edit or Remove; an edited row needs every
    one of its missing cells filled in with a valid value."""
    decisions, edits, row_errors, cell_errors = {}, {}, {}, {}
    canon_cache = {}

    for rid in ids:
        choice = posted.get(f"d_{rid}", "")
        if choice not in ("edit", "remove"):
            row_errors[rid] = "Choose Edit or Remove for this row."
            continue
        decisions[rid] = choice
        if choice != "edit":
            continue

        pos = df.index.get_loc(rid)
        for cidx in range(df.shape[1]):
            if not pd.isna(df.iat[pos, cidx]):
                continue
            value, error = _validate_manual_value(
                posted.get(f"v_{rid}_{cidx}", ""), kinds[cidx],
                df.iloc[:, cidx], canon_cache, cidx,
            )
            if error:
                cell_errors[(rid, cidx)] = error
            else:
                edits[(rid, cidx)] = value

    return decisions, edits, row_errors, cell_errors


def _short_value(value, limit=36):
    text = str(value)
    return text if len(text) <= limit else text[:limit - 1] + "…"


def _build_review_rows(df, ids, kinds, row_offset, posted=None, row_errors=None, cell_errors=None):
    """Everything the review template needs for each row with missing values:
    the cells to fill in, the rest of the row for context, and (when the page
    is re-shown after a mistake) what the user already chose and typed."""
    posted = posted or {}
    row_errors = row_errors or {}
    cell_errors = cell_errors or {}
    rows = []

    for rid in ids:
        pos = df.index.get_loc(rid)
        missing_cells, context_cells = [], []

        for cidx, name in enumerate(df.columns):
            value = df.iat[pos, cidx]
            if pd.isna(value):
                kind = kinds[cidx]
                missing_cells.append({
                    "idx": cidx,
                    "col": name,
                    "kind_label": KIND_LABELS[kind],
                    "input_type": KIND_INPUT_TYPES[kind],
                    "step": "1" if kind == "integer" else ("any" if kind == "number" else ""),
                    "value": posted.get(f"v_{rid}_{cidx}", ""),
                    "error": cell_errors.get((rid, cidx), ""),
                })
            else:
                context_cells.append({"col": name, "value": _short_value(value)})

        rows.append({
            "rid": rid,
            "label": rid + row_offset,
            "missing": missing_cells,
            "missing_names": ", ".join(c["col"] for c in missing_cells),
            "context": context_cells,
            "decision": posted.get(f"d_{rid}", ""),
            "error": row_errors.get(rid, ""),
        })

    return rows


def _render_review(request, context, df, missing, info, posted=None,
                   form_error="", row_errors=None, cell_errors=None):
    """The prompt shown right after 'Clean Data': how much is missing, and the
    choice between removing those rows and reviewing them by hand."""
    manual_allowed = len(missing["ids"]) <= MANUAL_REVIEW_MAX_ROWS
    posted = posted or {}

    context.update({
        "review": True,
        "missing": missing,
        "manual_allowed": manual_allowed,
        "manual_max": MANUAL_REVIEW_MAX_ROWS,
        "strategy": posted.get("strategy", ""),
        "form_error": form_error,
        "review_rows": (
            _build_review_rows(
                df, missing["ids"], _column_kinds(df), info["row_offset"],
                posted, row_errors, cell_errors,
            ) if manual_allowed else []
        ),
    })
    return render(request, "cascade_app/clean_data.html", context)


def _decision_summary(strategy, found, removed, rows_edited, cells_edited):
    """One plain-language line describing what the user chose."""
    if strategy == "none":
        return "No missing values were found, so no rows needed to be removed or edited."
    if strategy == "remove_all":
        return f"Removed all {_plural(removed, 'row')} that had missing values."
    parts = []
    if removed:
        parts.append(f"removed {_plural(removed, 'row')}")
    if rows_edited:
        parts.append(
            f"edited {_plural(rows_edited, 'row')} by typing in {_plural(cells_edited, 'missing value')}"
        )
    return (
        f"Reviewed {_plural(found, 'row')} with missing values manually: "
        + " and ".join(parts) + "."
    )


def _finish_cleaning(request, file_name, df, report, info, missing, strategy="none",
                     removed_ids=(), edit_log=(), rows_edited=0):
    """Final steps: report on what is left, save the cleaned file that the
    rest of the app reads, record the user's choices, then show the result."""
    df = df.reset_index(drop=True)
    rows_after = len(df)

    # Remaining missing values (none after the user's decisions)
    remaining_missing = int(df.isna().sum().sum())
    if remaining_missing > 0:
        report.append(
            f"{remaining_missing} missing value(s) remain and were not automatically filled."
        )

    # Flag possible outliers for manual review (never auto-changed)
    outlier_flags = _flag_outliers(df)

    if outlier_flags:
        flagged_desc = ", ".join(f"'{col}' ({count})" for col, count in outlier_flags)
        report.append(
            f"Flagged possible outliers for manual review in: {flagged_desc}. "
            f"Values were not changed or removed."
        )

    # Flag possible near-duplicate rows for manual review (never auto-removed)
    near_dupe_count = _flag_near_duplicate_rows(df)

    if near_dupe_count:
        report.append(
            f"Flagged {near_dupe_count} pair(s) of near-duplicate rows "
            f"(similar across most columns) for manual review. Rows were not merged or removed."
        )

    removed_ids = list(removed_ids)
    cells_edited = len(edit_log)
    summary = _decision_summary(
        strategy, missing["rows_with_missing"], len(removed_ids), rows_edited, cells_edited
    )

    # The user's choices, kept so Visualize and the PDF can show them.
    # (Lists are capped so the session stays small; the counts are exact.)
    actions = {
        "strategy": strategy,
        "summary": summary,
        "rows_with_missing": missing["rows_with_missing"],
        "rows_with_missing_pct": missing["rows_pct"],
        "cells_missing": missing["cells_missing"],
        "rows_removed": len(removed_ids),
        "rows_edited": rows_edited,
        "cells_edited": cells_edited,
        "rows_before": missing["total_rows"],
        "rows_after": rows_after,
        "removed_rows": [rid + info["row_offset"] for rid in removed_ids[:200]],
        "edits": [
            {"row": rid + info["row_offset"], "column": col, "value": str(value)}
            for rid, col, value in list(edit_log)[:100]
        ],
    }

    if strategy == "none":
        if not report:
            report.append("No cleaning was necessary. The dataset already appears clean.")
        else:
            report.append(summary)
    else:
        found = missing["rows_with_missing"]
        report.append(
            f"{_plural(found, 'row')} ({missing['rows_pct']} of the data) had missing values. "
            + summary
        )

    # Save the cleaned data as a real file on disk. Visualize, the PDF export
    # and the CSV download all read this one file, so every choice made above
    # is part of it.
    cleaned_filename = f"cleaned_{os.path.splitext(file_name)[0]}.csv"
    fs = FileSystemStorage()
    if fs.exists(cleaned_filename):
        fs.delete(cleaned_filename)
    fs.save(cleaned_filename, ContentFile(df.to_csv(index=False).encode("utf-8")))

    request.session["cleaned_file_name"] = cleaned_filename
    request.session["currency_symbols"] = info["currency_symbols"]
    request.session["cleaning_actions"] = actions
    request.session["cleaning_result"] = {
        "for_file": file_name,
        "cleaning_report": report,
        "original_rows": info["original_rows"],
        "cleaned_rows": rows_after,
        "original_columns": info["original_columns"],
        "cleaned_columns": len(df.columns),
        "remaining_missing": remaining_missing,
    }
    request.session.pop("cleaned_csv", None)
    # Chart/insight state from a previous cleaned file is now stale.
    for key in CLEANING_STATE_KEYS:
        request.session.pop(key, None)
    request.session.modified = True

    # Post/redirect/get: refreshing the report page must not re-run cleaning.
    return redirect("clean_data")


def _apply_decisions(request, context, file_name, df, report, info, missing):
    """Validate the user's choice and confirmation, apply it to the data and
    finish cleaning. If anything is wrong the review page is shown again with
    the user's input kept and the problems marked."""
    ids = missing["ids"]
    posted = request.POST
    strategy = posted.get("strategy", "")

    def back(message, row_errors=None, cell_errors=None):
        return _render_review(request, context, df, missing, info, posted,
                              message, row_errors, cell_errors)

    if strategy not in ("remove_all", "manual"):
        return back("Choose how to handle the rows with missing values.")
    if strategy == "manual" and len(ids) > MANUAL_REVIEW_MAX_ROWS:
        return back(
            f"Manual review is limited to {MANUAL_REVIEW_MAX_ROWS} rows at a time. "
            f"Remove those rows instead, or fix the missing values in your source file and upload it again."
        )

    removed_ids, edit_log, rows_edited = list(ids), [], 0

    if strategy == "manual":
        kinds = _column_kinds(df)
        decisions, edits, row_errors, cell_errors = _collect_manual_decisions(df, kinds, ids, posted)
        if row_errors or cell_errors:
            problems = len(set(row_errors) | {rid for rid, _ in cell_errors})
            return back(
                f"{_plural(problems, 'row')} still need attention. "
                f"Every row needs a decision, and every missing value in an edited row needs a valid entry.",
                row_errors, cell_errors,
            )
        removed_ids = [rid for rid, choice in decisions.items() if choice == "remove"]

    if posted.get("confirmed") != "1":
        return back("Please confirm your choices before they are applied.")

    if strategy == "manual":
        for (rid, cidx), value in edits.items():
            df.iat[df.index.get_loc(rid), cidx] = value
            edit_log.append((rid, df.columns[cidx], value))
        rows_edited = len({rid for rid, _ in edits})

    df = df.drop(index=removed_ids)
    if int(df.isna().sum().sum()) != 0:
        raise ValueError("Some missing values were left unresolved. Please try again.")

    return _finish_cleaning(request, file_name, df, report, info, missing, strategy,
                            removed_ids, edit_log, rows_edited)


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
            df, report, info = _prepare_dataset(file_path)
            missing = _missing_overview(df)

            # Nothing missing: nothing to decide, finish straight away.
            if not missing["ids"]:
                return _finish_cleaning(request, file_name, df, report, info, missing)

            # Step 1 ("Clean Data" clicked): show the numbers and ask.
            if request.POST.get("action") != "apply":
                return _render_review(request, context, df, missing, info)

            # Step 2 (decision submitted and confirmed): apply it.
            return _apply_decisions(request, context, file_name, df, report, info, missing)

        except Exception as e:
            context["error"] = f"Unable to clean file: {str(e)}"
    else:
        # Show the report of the cleaning that was last completed for this file.
        result = request.session.get("cleaning_result")
        if (
            result
            and result.get("for_file") == file_name
            and request.session.get("cleaned_file_name")
        ):
            context.update(result)
            context["cleaned"] = True
            context["cleaning_actions"] = request.session.get("cleaning_actions")

    return render(request, "cascade_app/clean_data.html", context)


@login_required(login_url='login')
def download_cleaned_data(request):
    # Serve the saved cleaned file, so the download always matches exactly what
    # Visualize uses (including the user's missing-value decisions).
    cleaned_name = request.session.get("cleaned_file_name")
    cleaned_path = os.path.join(settings.MEDIA_ROOT, cleaned_name) if cleaned_name else None

    if not cleaned_path or not os.path.exists(cleaned_path):
        return HttpResponse(
            "No cleaned file available.",
            status=404
        )

    with open(cleaned_path, "rb") as handle:
        content = handle.read()

    response = HttpResponse(
        content,
        content_type="text/csv"
    )

    response["Content-Disposition"] = (
        'attachment; filename="cascade_cleaned_data.csv"'
    )

    return response


# ---------------------------------------------------------------------------
# SAVE / RESTORE PROGRESS
#
# "Save Progress" downloads one .cascade file (a zip) that holds everything
# needed to carry on later:
#   manifest.json  the charts, the notes/insights the user edited, the Overall
#                  Summary and Recommendations text, filters, Key Metrics,
#                  pivot tables, their order, currency symbols and the record
#                  of cleaning choices
#   data.csv       the cleaned dataset those charts were built on
# Importing the file puts the dataset back and restores all of that exactly,
# so the user can open Visualize and keep adding to their insights without
# uploading or cleaning anything again. (Charts and the auto-written insights
# are rebuilt from the saved settings, so only the user's own edits need to be
# stored.)
#
# An imported file is untrusted input, so it is read with hard size limits,
# never extracted to disk, checked against its own checksum, and every saved
# setting is validated against the dataset before it is allowed into the
# session. Nothing is changed unless the whole file passes.
# ---------------------------------------------------------------------------

class ProgressImportError(Exception):
    """A saved-progress file that can't be used. The message is shown to the user."""


def _safe_ascii_name(text, fallback="visualization"):
    stem = os.path.splitext(os.path.basename(str(text or "")))[0]
    stem = re.sub(r"^cleaned_", "", stem)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("._-")
    return stem[:60] or fallback


def _download_progress(request, file_path):
    """Build the .cascade file for the visualization the user is looking at."""
    from django.utils import timezone

    with open(file_path, "rb") as handle:
        data = handle.read()
    df = read_dataset(file_path)

    state = {
        "charts": request.session.get("charts") or [],
        "insights_overrides": request.session.get("insights_overrides") or {},
        "summary_override": request.session.get("summary_override"),
        "recommendations_override": request.session.get("recommendations_override"),
        "filters": request.session.get("filters") or [],
        # None means "the default Key Metrics", an empty list means "none"
        "kpis": request.session.get("kpis"),
        "pivots": _session_pivots(request),
        "layout": _layout_for(request),
        "currency_symbols": request.session.get("currency_symbols") or {},
        "cleaning_actions": request.session.get("cleaning_actions"),
    }

    source = request.session.get("file_name") or request.session.get("cleaned_file_name") or ""
    now = timezone.now()
    manifest = {
        "format": PROGRESS_FORMAT,
        "version": PROGRESS_VERSION,
        "saved_at": now.isoformat(timespec="seconds"),
        "source_name": os.path.basename(str(source))[:120],
        "rows": int(len(df)),
        "columns": [str(c) for c in df.columns],
        "data_sha256": hashlib.sha256(data).hexdigest(),
        "state": state,
    }

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as bundle:
        bundle.writestr("manifest.json", json.dumps(manifest, ensure_ascii=False, indent=2))
        bundle.writestr("data.csv", data)

    filename = f"cascade-progress_{_safe_ascii_name(source or file_path)}_{now:%Y-%m-%d}{PROGRESS_EXTENSION}"
    response = HttpResponse(buffer.getvalue(), content_type="application/octet-stream")
    response["Content-Disposition"] = f'attachment; filename="{filename}"'
    return response


def _read_bundle_member(bundle, name, limit):
    """Read one file from the zip, never more than `limit` bytes (the declared
    size in a zip can be faked, so the real amount read is what is checked)."""
    try:
        with bundle.open(name) as member:
            content = member.read(limit + 1)
    except Exception:
        raise ProgressImportError(
            "That progress file looks damaged and could not be read. Try saving it again."
        )
    if len(content) > limit:
        raise ProgressImportError("That progress file is larger than Cascade can open.")
    return content


def _guarded(function, *args):
    """Run one of the validators; a malformed saved value just means that item
    is dropped instead of breaking the whole import."""
    try:
        return function(*args)
    except Exception:
        return None


def _clean_note(value, limit):
    if not isinstance(value, str):
        return None
    value = value.strip()
    return value[:limit] if value else None


def _sanitize_chart(df, raw):
    raw = raw if isinstance(raw, dict) else {}
    columns = set(df.columns)

    def column(key):
        value = raw.get(key, "")
        return value if isinstance(value, str) and value in columns else ""

    def calc(key):
        spec = _guarded(normalize_calc, df, raw.get(key))
        return dict(spec, on=True) if spec else None

    chart_type = raw.get("chart_type")
    period, agg = raw.get("period"), raw.get("agg")
    sort, topn = raw.get("sort"), raw.get("topn")
    chart = {
        "x_col": column("x_col"),
        "y_col": column("y_col"),
        "chart_type": chart_type if isinstance(chart_type, str) and (chart_type == "auto" or chart_type in CHART_TYPE_LABELS) else "auto",
        "period": period if isinstance(period, str) and period in TIME_PERIODS else "none",
        "agg": agg if isinstance(agg, str) and agg in TIME_AGGS else "sum",
        "calc": calc("calc"),
        "x_calc": calc("x_calc"),
        "sort": sort if isinstance(sort, str) and sort in SORT_OPTIONS else "default",
        "topn": topn if isinstance(topn, str) and topn in TOPN_MODES else "all",
        "topn_n": _rank_int(raw.get("topn_n")),
    }
    if chart["calc"] and chart["x_calc"]:
        chart["x_calc"] = None  # one calculation at a time, as in the editor
    return chart


def _sanitize_filter(df, raw):
    if isinstance(raw, dict) and isinstance(raw.get("values"), list):
        raw = dict(raw, values=raw["values"][:MAX_FILTER_CHOICES])
    return normalize_filter(df, raw)


def _sanitize_pivot(df, raw):
    raw = raw if isinstance(raw, dict) else {}
    cfg = raw.get("cfg")
    return {
        "cfg": _guarded(normalize_pivot_config, df, cfg) if isinstance(cfg, dict) else None,
        "insight": _clean_note(raw.get("insight"), MAX_NOTE_CHARS),
        "recommendation": _clean_note(raw.get("recommendation"), MAX_NOTE_CHARS),
    }


def _sanitize_cleaning_actions(raw):
    """Only the plain facts the page and PDF display, in the shape they expect."""
    if not isinstance(raw, dict) or raw.get("strategy") not in ("none", "remove_all", "manual"):
        return None

    def number(key):
        value = raw.get(key)
        return value if isinstance(value, int) and not isinstance(value, bool) and 0 <= value <= 10**9 else 0

    removed = [n for n in (raw.get("removed_rows") or [])[:200]
               if isinstance(n, int) and not isinstance(n, bool) and 0 <= n <= 10**9]
    edits = []
    for edit in (raw.get("edits") or [])[:100]:
        if isinstance(edit, dict) and isinstance(edit.get("row"), int) and not isinstance(edit.get("row"), bool):
            edits.append({
                "row": edit["row"],
                "column": str(edit.get("column", ""))[:80],
                "value": str(edit.get("value", ""))[:200],
            })
    return {
        "strategy": raw["strategy"],
        "summary": _clean_note(raw.get("summary"), 600) or "",
        "rows_with_missing": number("rows_with_missing"),
        "rows_with_missing_pct": (_clean_note(raw.get("rows_with_missing_pct"), 10) or ""),
        "cells_missing": number("cells_missing"),
        "rows_removed": number("rows_removed"),
        "rows_edited": number("rows_edited"),
        "cells_edited": number("cells_edited"),
        "rows_before": number("rows_before"),
        "rows_after": number("rows_after"),
        "removed_rows": removed,
        "edits": edits,
    }


def _read_progress_file(upload):
    """Open and fully validate a .cascade file. Returns (manifest, data_bytes,
    df, state) where `state` is the cleaned-up session state to restore."""
    if upload.size > MAX_PROGRESS_BYTES:
        raise ProgressImportError(
            f'That file is {_format_size(upload.size)}, which is over the {MAX_PROGRESS_MB} MB limit for progress files.'
        )

    wrong_file = "That is not a Cascade progress file. Use a file made with Save Progress (it ends in .cascade)."
    try:
        bundle = zipfile.ZipFile(io.BytesIO(upload.read()))
    except Exception:
        raise ProgressImportError(wrong_file)

    with bundle:
        infos = bundle.infolist()
        if len(infos) != 2 or {info.filename for info in infos} != {"manifest.json", "data.csv"}:
            raise ProgressImportError(wrong_file)
        manifest_raw = _read_bundle_member(bundle, "manifest.json", MAX_PROGRESS_MANIFEST_BYTES)
        data = _read_bundle_member(bundle, "data.csv", MAX_PROGRESS_DATA_BYTES)

    try:
        manifest = json.loads(manifest_raw.decode("utf-8"))
    except (UnicodeDecodeError, ValueError, RecursionError):
        raise ProgressImportError(wrong_file)
    if not isinstance(manifest, dict) or manifest.get("format") != PROGRESS_FORMAT:
        raise ProgressImportError(wrong_file)

    version = manifest.get("version")
    if not isinstance(version, int) or isinstance(version, bool) or version < 1:
        raise ProgressImportError(wrong_file)
    if version > PROGRESS_VERSION:
        raise ProgressImportError(
            "That progress file was saved by a newer version of Cascade than this one, so it can't be opened here."
        )

    if manifest.get("data_sha256") != hashlib.sha256(data).hexdigest():
        raise ProgressImportError(
            "That progress file is damaged or was changed after it was saved (its data no longer matches). "
            "Save it again from Visualize and use the new file."
        )

    try:
        df = pd.read_csv(io.BytesIO(data))
    except Exception:
        raise ProgressImportError("The data inside that progress file could not be read.")
    if df.shape[1] == 0 or df.empty:
        raise ProgressImportError("That progress file contains no data.")
    if df.shape[1] > MAX_PROGRESS_COLUMNS:
        raise ProgressImportError("The data inside that progress file has too many columns.")
    if [str(c) for c in df.columns] != manifest.get("columns"):
        raise ProgressImportError("That progress file is inconsistent (its columns do not match its data).")

    raw = manifest.get("state")
    raw = raw if isinstance(raw, dict) else {}

    # ---- charts, the notes the user wrote on them, and the page order ----
    raw_charts = raw.get("charts") if isinstance(raw.get("charts"), list) else []
    charts = [_sanitize_chart(df, c) for c in raw_charts[:MAX_SAVED_CHARTS]]
    if not charts:
        charts = [{"x_col": "", "y_col": "", "chart_type": "auto"}]

    overrides = {}
    if isinstance(raw.get("insights_overrides"), dict):
        for key, text in raw["insights_overrides"].items():
            if isinstance(key, str) and key.isascii() and key.isdigit() and len(key) <= 4 and int(key) < len(charts):
                note = _clean_note(text, MAX_NOTE_CHARS)
                if note:
                    overrides[str(int(key))] = note

    raw_pivots = raw.get("pivots")
    if isinstance(raw_pivots, list):
        pivots = [_sanitize_pivot(df, p) for p in raw_pivots[:MAX_SAVED_PIVOTS]]
    else:
        pivots = [{"cfg": None, "insight": None, "recommendation": None}]

    layout = [t for t in raw.get("layout", []) if isinstance(t, str)] if isinstance(raw.get("layout"), list) else []

    # ---- filters and Key Metrics ----
    filters = []
    if isinstance(raw.get("filters"), list):
        for item in raw["filters"][:MAX_FILTERS * 2]:
            spec = _guarded(_sanitize_filter, df, item)
            if spec and all(spec["col"] != f["col"] for f in filters) and len(filters) < MAX_FILTERS:
                filters.append(spec)

    kpis = None
    if isinstance(raw.get("kpis"), list):
        kpis = []
        for item in raw["kpis"][:MAX_KPIS * 2]:
            spec = _guarded(normalize_kpi, df, item)
            if spec and spec not in kpis and len(kpis) < MAX_KPIS:
                kpis.append(spec)

    # ---- written summary, recommendations, currency, cleaning record ----
    recommendations = None
    if isinstance(raw.get("recommendations_override"), list):
        lines = [_clean_note(line, 1000) for line in raw["recommendations_override"][:MAX_RECOMMENDATIONS]]
        recommendations = [line for line in lines if line] or None

    currency = {}
    if isinstance(raw.get("currency_symbols"), dict):
        for col, symbol in raw["currency_symbols"].items():
            if isinstance(col, str) and col in df.columns and isinstance(symbol, str) and 0 < len(symbol) <= 8:
                currency[col] = symbol

    state = {
        "charts": charts,
        "insights_overrides": overrides,
        "summary_override": _clean_note(raw.get("summary_override"), MAX_SUMMARY_CHARS),
        "recommendations_override": recommendations,
        "filters": filters,
        "kpis": kpis,
        "pivots": pivots,
        "layout": _normalize_layout(layout, len(charts), len(pivots)),
        "currency_symbols": currency,
        "cleaning_actions": _guarded(_sanitize_cleaning_actions, raw.get("cleaning_actions")),
    }
    return manifest, data, df, state


def _import_progress(request, upload):
    """Restore a saved visualization. The current work is only replaced once
    the whole file has been read and checked."""
    try:
        manifest, data, df, state = _read_progress_file(upload)
    except ProgressImportError as error:
        return _render_upload(request, import_error=str(error))
    except Exception:
        return _render_upload(
            request, import_error="That progress file could not be opened. It may be damaged."
        )

    # The dataset goes in as a new file: a name that is already taken gets a
    # unique suffix, so nobody's existing file is overwritten.
    stem = _safe_ascii_name(manifest.get("source_name"), "progress")
    saved_name = FileSystemStorage().save(f"cleaned_{stem}.csv", ContentFile(data))

    # Forget whatever was open, then restore the saved state.
    for key in ("file_name", "cleaned_csv", "cleaning_result", "cleaning_actions") + CLEANING_STATE_KEYS:
        request.session.pop(key, None)

    request.session["cleaned_file_name"] = saved_name
    request.session["currency_symbols"] = state["currency_symbols"]
    if state["cleaning_actions"]:
        request.session["cleaning_actions"] = state["cleaning_actions"]
    request.session["charts"] = state["charts"]
    request.session["insights_overrides"] = state["insights_overrides"]
    request.session["summary_override"] = state["summary_override"]
    request.session["recommendations_override"] = state["recommendations_override"]
    request.session["filters"] = state["filters"]
    if state["kpis"] is not None:
        request.session["kpis"] = state["kpis"]
    request.session["pivots"] = state["pivots"]
    request.session["layout"] = state["layout"]

    saved_on = str(manifest.get("saved_at") or "")[:10]
    request.session["progress_notice"] = (
        f"Restored your saved progress{f' from {saved_on}' if saved_on else ''}: "
        f"{_plural(len(state['charts']), 'chart')}, {_plural(len(state['pivots']), 'pivot table')} "
        f"and {_plural(len(df), 'row')} of data."
    )
    request.session.modified = True
    return redirect("visualize_data")