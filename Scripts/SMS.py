def run_pipeline(tp_df):
    concat_file = build_concat()

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    tp_output = f"TP_SMS_{ts}.xlsx"

    sheet_map = create_tp_sheets(tp_df, tp_output)

    route_sms(concat_file, tp_df, tp_output)

    return tp_output