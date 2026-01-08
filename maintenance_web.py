from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, send_file

from maintenance_alert import (
    Config,
    DualLogger,
    ExcelHandler,
    MaintenanceChecker,
    StatisticsManager,
    ReportGenerator,
    EmailSender,
)

app = Flask(__name__)

# Core objects reused from maintenance_alert
config = Config()
logger = DualLogger(config.LOG_FILE)
excel_handler = ExcelHandler(config, logger)
maintenance_checker = MaintenanceChecker(config, logger)
statistics_manager = StatisticsManager(config, logger)
report_generator = ReportGenerator(config, logger, maintenance_checker, statistics_manager)
email_sender = EmailSender(config, logger)


def _build_items_list(dfs, status_label: str):
    items = []
    for df in dfs:
        for _, row in df.iterrows():
            item_type = row.get("Тип", "")
            designation = row.get("Обозначение", "")
            object_name = row.get("Объект", "")
            items.append(
                {
                    "type": item_type,
                    "status": status_label,
                    "designation": designation,
                    "object": object_name,
                    "html": maintenance_checker.format_item_info(row, item_type),
                }
            )
    return items


@app.route("/")
def dashboard():
    sheet_type = request.args.get("sheet_type", "all")
    status_filter = request.args.get("status", "all")
    designation_filter = request.args.get("designation", "").strip()
    object_filter = request.args.get("object", "all")
    email_status = request.args.get("email_status")
    chart_date = request.args.get("chart_date", "").strip()
    serviced_status = request.args.get("serviced_status")
    serviced_message = request.args.get("serviced_message", "")
    
    # Calculate offset from chart_date
    chart_offset = 0
    if chart_date:
        try:
            from datetime import datetime
            selected_date = datetime.strptime(chart_date, "%Y-%m-%d").date()
            today = datetime.now().date()
            chart_offset = (selected_date - today).days
        except ValueError:
            chart_date = ""
            chart_offset = 0
    
    # If no date specified, use today's date
    if not chart_date:
        from datetime import datetime
        chart_date = datetime.now().strftime("%Y-%m-%d")

    urgent_items, warning_items, total_records, status_counts, recalc_success = excel_handler.read_data()

    urgent_list = _build_items_list(urgent_items, "urgent")
    warning_list = _build_items_list(warning_items, "warning")

    # Collect unique objects for dropdown
    all_items = urgent_list + warning_list
    unique_objects = sorted(set(item.get("object", "") for item in all_items if item.get("object")))

    def apply_filters(items):
        if sheet_type != "all":
            items = [i for i in items if i["type"] == sheet_type]
        if designation_filter:
            # Filter by designation - case insensitive substring match
            items = [i for i in items if designation_filter.lower() in str(i.get("designation", "")).lower()]
        if object_filter != "all":
            items = [i for i in items if i.get("object") == object_filter]
        return items

    show_urgent = status_filter in ("all", "urgent")
    show_warning = status_filter in ("all", "warning")

    filtered_urgent = apply_filters(urgent_list) if show_urgent else []
    filtered_warning = apply_filters(warning_list) if show_warning else []

    filtered_urgent_count = len(filtered_urgent)
    filtered_warning_count = len(filtered_warning)
    total_urgent = len(urgent_list)
    total_warning = len(warning_list)

    # Use filtered counts for percentage if filters are applied? 
    # Actually, keep the global stats but maybe show filtered ones.
    unserviced_count = status_counts.get(config.STATUS_URGENT, 0)
    unserviced_percentage = (unserviced_count / total_records * 100) if total_records else 0.0

    return render_template(
        "dashboard.html",
        config=config,
        status_counts=status_counts,
        total_records=total_records,
        unserviced_percentage=unserviced_percentage,
        urgent_items=filtered_urgent,
        warning_items=filtered_warning,
        total_urgent=total_urgent,
        total_warning=total_warning,
        filtered_urgent_count=filtered_urgent_count,
        filtered_warning_count=filtered_warning_count,
        sheet_type=sheet_type,
        status_filter=status_filter,
        designation_filter=designation_filter,
        object_filter=object_filter,
        unique_objects=unique_objects,
        chart_offset=chart_offset,
        chart_date=chart_date,
        email_status=email_status,
        recalc_success=recalc_success,
        serviced_status=serviced_status,
        serviced_message=serviced_message,
    )


@app.route("/stats")
def stats():
    stats_data = statistics_manager.get_statistics()
    return render_template("stats.html", config=config, stats=stats_data)


@app.route("/settings")
def settings():
    return render_template("settings.html", config=config)


@app.route("/chart.png")
def chart_png():
    offset_days = request.args.get("offset", "0")
    try:
        offset_days = int(offset_days)
    except ValueError:
        offset_days = 0
    
    chart_path = statistics_manager.create_chart(offset_days=offset_days)
    if not chart_path or not Path(chart_path).exists():
        return ("Диаграмма недоступна", 404)
    return send_file(chart_path, mimetype="image/png")


@app.route("/send-email", methods=["POST"])
def send_email():
    urgent_items, warning_items, total_records, status_counts, recalc_success = excel_handler.read_data()

    total_alarm = sum(len(df) for df in urgent_items) if urgent_items else 0
    total_warning = sum(len(df) for df in warning_items) if warning_items else 0

    if total_alarm == 0 and total_warning == 0:
        return redirect(url_for("dashboard", email_status="no_items"))

    statistics_manager.update_statistics(urgent_items, warning_items, total_records, status_counts)

    email_body, chart_path = report_generator.create_body(
        urgent_items,
        warning_items,
        total_records,
        status_counts,
        recalc_success,
    )

    maintenance_data_file = None
    if urgent_items:
        maintenance_data_file = excel_handler.generate_maintenance_data_file(urgent_items)

    sent = email_sender.send(email_body, config.RECIPIENTS, chart_path, maintenance_data_file)

    if sent and maintenance_data_file and maintenance_data_file.exists():
        try:
            maintenance_data_file.unlink()
        except Exception:
            pass

    status = "sent" if sent else "error"
    return redirect(url_for("dashboard", email_status=status))


@app.route("/download-excel")
def download_excel():
    # File is usually in TMP_DIR after recalculation
    file_path = config.TMP_DIR / config.EXCEL_FILENAME
    if not file_path.exists():
        # Fallback to original file if tmp doesn't exist
        file_path = config.get_excel_file_path()
        
    if not file_path.exists():
        return ("Файл не найден", 404)
        
    return send_file(
        file_path,
        as_attachment=True,
        download_name=config.EXCEL_FILENAME
    )


@app.route("/mark-serviced", methods=["POST"])
def mark_serviced():
    """
    Отмечает оборудование как обслуженное, обновляя дату последнего ТО в Excel файле.
    Принимает параметры: sheet_name (название листа) и designation (обозначение оборудования)
    """
    sheet_name = request.form.get("sheet_name", "").strip()
    designation = request.form.get("designation", "").strip()
    
    # Preserve current filters
    object_filter = request.form.get("object", "all")
    status_filter = request.form.get("status", "all")
    designation_filter = request.form.get("designation_filter", "")
    
    if not sheet_name or not designation:
        return redirect(url_for("dashboard", 
                               object=object_filter,
                               status=status_filter, 
                               designation=designation_filter,
                               serviced_status="missing_params"))
    
    success, message = excel_handler.mark_as_serviced(sheet_name, designation)
    
    status_param = "success" if success else "error"
    
    return redirect(url_for("dashboard", 
                           object=object_filter,
                           status=status_filter,
                           designation=designation_filter,
                           serviced_status=status_param,
                           serviced_message=message))


@app.route("/mark-bulk-serviced", methods=["POST"])
def mark_bulk_serviced():
    """
    Отмечает несколько единиц оборудования как обслуженное.
    Принимает JSON-массив items с объектами {sheet_name, designation}
    """
    import json
    
    items_json = request.form.get("items", "[]")
    object_filter = request.form.get("object", "all")
    status_filter = request.form.get("status", "all")
    designation_filter = request.form.get("designation_filter", "")
    
    try:
        items = json.loads(items_json)
    except json.JSONDecodeError:
        return redirect(url_for("dashboard",
                               object=object_filter,
                               status=status_filter,
                               designation=designation_filter,
                               serviced_status="error",
                               serviced_message="Ошибка обработки данных"))
    
    if not items:
        return redirect(url_for("dashboard",
                               object=object_filter,
                               status=status_filter,
                               designation=designation_filter,
                               serviced_status="error",
                               serviced_message="Не выбрано оборудование"))
    
    success_count = 0
    error_count = 0
    errors = []
    
    for item in items:
        sheet_name = item.get("sheet_name", "").strip()
        designation = item.get("designation", "").strip()
        
        if not sheet_name or not designation:
            error_count += 1
            errors.append(f"Пропущено: неполные данные")
            continue
        
        success, message = excel_handler.mark_as_serviced(sheet_name, designation)
        
        if success:
            success_count += 1
        else:
            error_count += 1
            errors.append(f"{designation}: {message}")
    
    # Prepare result message
    if error_count == 0:
        status_param = "success"
        result_message = f"Успешно обслужено: {success_count} ед. оборудования"
    elif success_count == 0:
        status_param = "error"
        result_message = f"Ошибка обслуживания всех {error_count} ед. оборудования. " + "; ".join(errors[:3])
    else:
        status_param = "success"
        result_message = f"Обслужено: {success_count} ед., ошибок: {error_count} ед."
    
    return redirect(url_for("dashboard",
                           object=object_filter,
                           status=status_filter,
                           designation=designation_filter,
                           serviced_status=status_param,
                           serviced_message=result_message))


if __name__ == "__main__":
    # For production you will likely set host/port and disable debug,
    # but this is fine for local testing.
    app.run(debug=True)
