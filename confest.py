import os
import time
import glob
import pytest
from datetime import datetime

# Where reports will be saved
REPORT_DIR = r"C:\1040ta5\Reports"

def pytest_configure(config):
    """Automatically forces VS Code to generate HTML and Excel reports with timestamps."""
    if not os.path.exists(REPORT_DIR):
        os.makedirs(REPORT_DIR)
        
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    
    # Force HTML and Excel generation even when running from VS Code UI
    config.option.htmlpath = os.path.join(REPORT_DIR, f"FADS_Report_{timestamp}.html")
    config.option.self_contained_html = True
    config.option.excelreport = os.path.join(REPORT_DIR, f"FADS_Report_{timestamp}.xlsx")

@pytest.fixture(scope="session", autouse=True)
def cleanup_old_reports():
    """Deletes reports older than 24 hours before VS Code starts the tests."""
    print("\n[🧹] Running pre-test cleanup...")
    now = time.time()
    for ext in ["*.html", "*.xlsx"]:
        for file in glob.glob(os.path.join(REPORT_DIR, ext)):
            if (now - os.path.getmtime(file)) > 86400: # 86400 seconds = 24 hours
                try: os.remove(file)
                except: pass
    yield

@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item, call):
    """Captures exact Start Time, End Time, and Duration for the HTML report."""
    outcome = yield
    report = outcome.get_result()
    if report.when == 'call':
        start_dt = datetime.fromtimestamp(call.start)
        stop_dt = datetime.fromtimestamp(call.stop)
        report.custom_start_time = start_dt.strftime('%H:%M:%S')
        report.custom_end_time = stop_dt.strftime('%H:%M:%S')
        report.custom_duration = f"{(stop_dt - start_dt).total_seconds():.2f}s"

def pytest_html_results_table_header(cells):
    """Injects custom column headers into the HTML report."""
    cells.insert(2, '<th class="sortable text-center">Start Time</th>')
    cells.insert(3, '<th class="sortable text-center">End Time</th>')
    cells.insert(4, '<th class="sortable text-center">Duration</th>')

def pytest_html_results_table_row(report, cells):
    """Injects custom timing data into the HTML report rows."""
    cells.insert(2, f'<td class="text-center">{getattr(report, "custom_start_time", "")}</td>')
    cells.insert(3, f'<td class="text-center">{getattr(report, "custom_end_time", "")}</td>')
    cells.insert(4, f'<td class="text-center">{getattr(report, "custom_duration", "")}</td>')
