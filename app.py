"""
Snatched Beauty Box — Web service wrapper for recommend_v5.py

This does NOT contain any recommendation logic itself — it just receives
requests from Make.com, saves the uploaded spreadsheet, calls the existing
(already tested) recommend_v5.py functions, and returns the result as JSON.

Endpoints:
  POST /recommend
    form-data:
      file: the current .xlsx spreadsheet (Make.com exports this fresh
            from Google Sheets before each call, so the engine always
            works with real, current data)
      customer_name: e.g. "Sandra Hoffmann"
      target_month: optional, e.g. "2026-09" — omit to use today's date
      box_type_override: optional, e.g. "Essentials"
    returns: JSON with the recommended products, value summary, warnings

  POST /swap
    form-data:
      file: the current .xlsx spreadsheet
      customer_name, category, tier: which slot is being swapped
      already_rejected: comma-separated list of product names already tried
      current_box_products: comma-separated list of the other 4 products
                             already in the box (avoids same-brand repeats)
    returns: JSON with the next alternative, or a clear "none left" message
"""
from flask import Flask, request, jsonify
import os
import tempfile
import importlib
import recommend_v5

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25MB — the exported
                                                        # spreadsheet with all
                                                        # sheets can be a few MB;
                                                        # Flask's default limit
                                                        # is much smaller and was
                                                        # silently rejecting it
                                                        # (413 error) before this.

@app.route('/', methods=['GET'])
def health_check():
    # Simple endpoint to confirm the service is alive — Make.com or a
    # browser can hit this to check the deployment worked.
    return jsonify({'status': 'ok', 'service': 'snatched-recommend-engine'})

@app.route('/recommend', methods=['POST'])
def recommend_endpoint():
    if 'file' not in request.files:
        return jsonify({'error': 'No spreadsheet file was sent.'}), 400
    customer_name = request.form.get('customer_name')
    if not customer_name:
        return jsonify({'error': 'customer_name is required.'}), 400

    target_month = request.form.get('target_month') or None
    box_type_override = request.form.get('box_type_override') or None

    uploaded = request.files['file']
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        importlib.reload(recommend_v5)
        recommend_v5.set_workbook_path(tmp_path)
        recommend_v5.set_target_month(target_month)

        out = recommend_v5.recommend(customer_name, box_type_override=box_type_override)
        if isinstance(out, str):
            return jsonify({'error': out}), 404

        (cust, picks, warnings, received, ratio, recent, hard_block, soft_avoid,
         hist_pat, total, timeline, inv_cat_map, value_summary) = out

        return jsonify({
            'customer_name': cust['name'],
            'box_type': cust['box_type'],
            'products': [
                {
                    'name': p['name'],
                    'category': p['category'],
                    'tier': p['tier'],
                    'stock': p['stock'],
                    'price_aed': p.get('retail_price_aed'),
                }
                for p in picks
            ],
            'value_summary': value_summary,
            'warnings': warnings,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_path)

@app.route('/swap', methods=['POST'])
def swap_endpoint():
    """Simplified — Make.com only needs to tell us WHICH ROW changed in
    pending_approval. We figure out the customer/category/tier ourselves
    by reading that row directly, instead of asking Make.com to extract
    those pieces separately (which its Watch Changes trigger can't cleanly do)."""
    row_number = request.form.get('row_number')
    if not row_number:
        return jsonify({'error': 'row_number is required.'}), 400
    row_number = int(float(row_number))

    sheet_name = request.form.get('sheet_name', 'pending_approval')
    already_rejected = request.form.get('already_rejected', '')
    already_rejected = [s.strip() for s in already_rejected.split(',') if s.strip()]
    target_month = request.form.get('target_month') or None

    try:
        importlib.reload(recommend_v5)
        recommend_v5.set_target_month(target_month)

        ws = recommend_v5.wb[sheet_name]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        name_col = headers.index('Product Name') + 1
        cat_col = headers.index('Category') + 1
        tier_col = headers.index('Tier') + 1

        category = ws.cell(row_number, cat_col).value
        tier = ws.cell(row_number, tier_col).value
        current_product = ws.cell(row_number, name_col).value

        # The product currently sitting in this slot must never be
        # recommended back to itself as the "next alternative".
        if current_product and current_product not in already_rejected:
            already_rejected = already_rejected + [current_product]

        # Find which customer this row belongs to by scanning upward for
        # the nearest customer-header row above it.
        customer_name = None
        for r in range(row_number, 0, -1):
            fc = ws.cell(r, 1).value
            if isinstance(fc, str) and '|' in fc and 'Month:' in fc:
                customer_name = fc.split('|')[0].strip()
                break
        if not customer_name:
            return jsonify({'error': f'Could not find which customer owns row {row_number}.'}), 400

        # Also collect the other 4 products already in this customer's box,
        # so we don't recommend the same brand twice.
        current_box_products = []
        r = row_number
        while True:
            n = ws.cell(r, headers.index('#') + 1).value
            if not isinstance(n, (int, float)):
                break
            prod = ws.cell(r, name_col).value
            if r != row_number and prod:
                current_box_products.append(prod)
            r -= 1
        r = row_number + 1
        while True:
            n = ws.cell(r, headers.index('#') + 1).value
            if not isinstance(n, (int, float)):
                break
            prod = ws.cell(r, name_col).value
            if prod:
                current_box_products.append(prod)
            r += 1

        product, message = recommend_v5.get_next_alternative(
            customer_name, category, tier, already_rejected, current_box_products
        )
        if product is None:
            return jsonify({'found': False, 'message': message})

        return jsonify({
            'found': True,
            'product': {
                'name': product['name'],
                'category': product['category'],
                'tier': product['tier'],
                'stock': product['stock'],
                'price_aed': product.get('retail_price_aed'),
            },
            'message': message,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/generate_month', methods=['POST'])
def generate_month_endpoint():
    """Generates a full month's recommendations for EVERY active/due customer
    in ONE call — not one customer at a time. This matters: it shares stock
    tracking across all customers in the same call, so two different
    customers can never accidentally get recommended the same last-1-in-stock
    item within the same month's batch."""
    if 'file' not in request.files:
        return jsonify({'error': 'No spreadsheet file was sent.'}), 400
    target_month = request.form.get('target_month')
    if not target_month:
        return jsonify({'error': 'target_month is required, e.g. 2026-09'}), 400

    uploaded = request.files['file']
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        importlib.reload(recommend_v5)
        recommend_v5.set_workbook_path(tmp_path)
        recommend_v5.set_target_month(target_month)
        recommend_v5.reset_allocations()

        customers = recommend_v5.load_customers()
        all_rec, recent_boxes, box_timeline = recommend_v5.load_box_history()

        results = []
        for cid, cust in customers.items():
            due, reason = recommend_v5.is_customer_due(cust, box_timeline, target_month)
            if not due:
                continue
            out = recommend_v5.recommend(cust['name'])
            if isinstance(out, str):
                continue
            (c, picks, warnings, received, ratio, recent, hard_block, soft_avoid,
             hist_pat, total, timeline, inv_cat_map, value_summary) = out
            results.append({
                'customer_name': c['name'],
                'customer_id': c['id'],
                'box_type': c['box_type'],
                'products': [
                    {'name': p['name'], 'category': p['category'], 'tier': p['tier'],
                     'stock': p['stock'], 'price_aed': p.get('retail_price_aed')}
                    for p in picks
                ],
                'value_summary': value_summary,
                'warnings': warnings,
            })

        return jsonify({'target_month': target_month, 'customers': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_path)

@app.route('/generate_month_formatted', methods=['POST'])
def generate_month_formatted_endpoint():
    """Does the whole job in one call: generates recommendations for every
    due customer AND writes them into a properly formatted sheet (pink
    headers, purple customer rows — same style as every month built so
    far), reusing the same tested build_month_sheet() function. Returns
    the ready-to-use file — Make.com just uploads it back to replace the
    live Google Sheet, no Iterator or row-by-row writing needed at all."""
    if 'file' not in request.files:
        return jsonify({'error': 'No spreadsheet file was sent.'}), 400
    target_month = request.form.get('target_month')
    sheet_name = request.form.get('sheet_name')
    if not target_month:
        return jsonify({'error': 'target_month is required, e.g. 2026-10'}), 400
    if not sheet_name:
        sheet_name = f'pending_approval_{target_month.replace("-", "_")}'

    uploaded = request.files['file']
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        import build_final_sheet
        importlib.reload(recommend_v5)
        importlib.reload(build_final_sheet)
        recommend_v5.set_workbook_path(tmp_path)
        recommend_v5.set_target_month(target_month)
        build_final_sheet.set_sheet_path(tmp_path)

        build_final_sheet.build_month_sheet(target_month, sheet_name)

        from flask import send_file
        return send_file(tmp_path, as_attachment=True,
                          download_name='snatched_beauty_box_master_updated.xlsx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/commit_month', methods=['POST'])
def commit_month_endpoint():
    """Called when the owner clicks 'Approve & Lock In' for a month.
    Reads the CURRENT state of the spreadsheet directly — whatever she's
    already approved/swapped in the pending_approval sheet for that month —
    and commits exactly those products. Make.com's job is simple: export
    the current sheet, send it here, upload back whatever comes back.
    No complex data-building required on the Make.com side at all."""
    if 'file' not in request.files:
        return jsonify({'error': 'No spreadsheet file was sent.'}), 400
    target_month = request.form.get('target_month')
    sheet_name = request.form.get('sheet_name', 'pending_approval')
    if not target_month:
        return jsonify({'error': 'target_month is required.'}), 400

    uploaded = request.files['file']
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        uploaded.save(tmp.name)
        tmp_path = tmp.name

    try:
        import openpyxl as oxl
        from collections import Counter, defaultdict
        wb = oxl.load_workbook(tmp_path)

        ws_pa = wb[sheet_name]
        headers = [ws_pa.cell(1, c).value for c in range(1, ws_pa.max_column + 1)]
        num_col = headers.index('#') + 1
        name_col = headers.index('Product Name') + 1
        status_col = headers.index('STATUS ✏️') + 1

        current_customer = None
        approved = defaultdict(list)
        for row in ws_pa.iter_rows(min_row=1, values_only=True):
            fc = row[0]
            if isinstance(fc, str) and '|' in fc and 'Month:' in fc:
                current_customer = fc.split('|')[0].strip()
                continue
            n = row[num_col-1]
            if isinstance(n, (int, float)) and row[status_col-1] == 'approved':
                approved[current_customer].append(row[name_col-1])

        wsc = wb['customers']
        hc = [wsc.cell(1, c).value for c in range(1, wsc.max_column + 1)]
        name_col_c = hc.index('name') + 1
        cid_col = hc.index('customer_id') + 1
        box_col = hc.index('box_type') + 1
        name_to_id, name_to_boxtype = {}, {}
        for r in wsc.iter_rows(min_row=2, values_only=True):
            if r[name_col_c-1]:
                name_to_id[r[name_col_c-1]] = r[cid_col-1]
                name_to_boxtype[r[name_col_c-1]] = r[box_col-1]

        ws_boxes = wb['boxes']
        ws_items = wb['box_items']
        next_box_num = ws_boxes.max_row + 1
        next_item_num = ws_items.max_row

        usage = Counter()
        committed_customers = []
        for cust_name, products in approved.items():
            cid = name_to_id.get(cust_name)
            if not cid or not products:
                continue
            box_id = f'BOX{next_box_num:04d}'
            next_box_num += 1
            box_type = name_to_boxtype.get(cust_name)
            ws_boxes.append([box_id, cid, cust_name, target_month, 'sent', box_type])
            for prod in products:
                next_item_num += 1
                ws_items.append([f'BI{next_item_num:05d}', cust_name, cid, target_month, box_type, box_id, prod])
                usage[prod] += 1
            committed_customers.append(cust_name)

        ws_inv = wb['inventory']
        hi = [ws_inv.cell(1, c).value for c in range(1, ws_inv.max_column + 1)]
        n_col = hi.index('name') + 1
        s_col = hi.index('stock_qty') + 1
        for row in ws_inv.iter_rows(min_row=2):
            name = row[n_col-1].value
            if name in usage:
                old = row[s_col-1].value or 0
                row[s_col-1].value = max(0, old - usage[name])

        wb.save(tmp_path)

        from flask import send_file
        return send_file(tmp_path, as_attachment=True,
                          download_name='snatched_beauty_box_master_updated.xlsx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
