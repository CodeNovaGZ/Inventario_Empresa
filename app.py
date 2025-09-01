from flask import Flask, render_template, request, redirect, url_for, session, flash
from werkzeug.security import check_password_hash
from datetime import datetime
import excel_utils as utils

app = Flask(__name__)
app.secret_key = 'cambia_esto_por_una_clave_secreta'

utils.ensure_files()

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('user_id') or not session.get('is_admin'):
            flash('Se requiere acceso de administrador')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

@app.route('/')
@login_required
def index():
    products = utils.load_products()
    return render_template('index.html', products=products)

@app.route('/login', methods=['GET','POST'])
def login():
    if session.get('user_id'):
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        users = utils.load_users()
        user = next((u for u in users if u['username'] == username), None)
        if user and check_password_hash(user['password_hash'], password):
            session['user_id'] = user['id']
            session['username'] = user['username']
            session['is_admin'] = bool(user.get('is_admin'))
            flash('Bienvenido!')
            return redirect(url_for('index'))
        else:
            flash('Credenciales inválidas')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/admin')
@admin_required
def admin():
    logs = utils.load_logs(200); products = utils.load_products()
    prod_map = {p['id']: p['name'] for p in products}; enriched_logs=[]
    for l in logs: enriched_logs.append({'id':l['id'],'product_name':prod_map.get(l['product_id'],'—'),'change_amount':l['change_amount'],'reason':l['reason'],'username':l['username'],'timestamp':l['timestamp']})
    orders = utils.load_orders(100)
    return render_template('admin.html', logs=enriched_logs, orders=orders)

@app.route('/create_product', methods=['POST'])
@admin_required
def create_product():
    name = request.form.get('name'); category = request.form.get('category'); model = request.form.get('model')
    color = request.form.get('color'); size = request.form.get('size')
    try: price = float(request.form.get('price') or 0)
    except: price = 0.0
    try: stock = int(request.form.get('stock') or 0)
    except: stock = 0
    created_at = datetime.utcnow().isoformat(); pid = utils._next_id_from_sheet(utils.PRODUCTS_XLSX,'products')
    product = {'id':pid,'name':name,'category':category,'model':model,'color':color,'size':size,'price':price,'stock':stock,'created_at':created_at}
    utils.save_product_row(product); utils.append_log(pid, stock, 'stock inicial', session.get('username')); flash('Producto creado')
    return redirect(url_for('admin'))

@app.route('/edit_product/<int:product_id>', methods=['GET','POST'])
@admin_required
def edit_product(product_id):
    products = utils.load_products(); p = next((x for x in products if x['id']==product_id), None)
    if not p: flash('Producto no encontrado'); return redirect(url_for('index'))
    if request.method == 'POST':
        p['name'] = request.form.get('name'); p['category'] = request.form.get('category'); p['model'] = request.form.get('model')
        p['color'] = request.form.get('color'); p['size'] = request.form.get('size')
        try: p['price'] = float(request.form.get('price') or 0)
        except: p['price'] = 0
        try: p['stock'] = int(request.form.get('stock') or 0)
        except: p['stock'] = 0
        utils.save_product_row(p); flash('Producto actualizado'); return redirect(url_for('index'))
    return render_template('edit_product.html', p=p)

@app.route('/delete_product/<int:product_id>', methods=['POST'])
@admin_required
def delete_product(product_id):
    utils.delete_product_by_id(product_id); flash('Producto eliminado'); return redirect(url_for('index'))

@app.route('/adjust_stock/<int:product_id>', methods=['POST'])
@admin_required
def adjust_stock(product_id):
    try: change = int(request.form.get('change') or 0)
    except: change = 0
    reason = request.form.get('reason') or ''
    products = utils.load_products(); p = next((x for x in products if x['id']==product_id), None)
    if not p: flash('Producto no existe'); return redirect(url_for('index'))
    p['stock'] = p['stock'] + change; utils.save_product_row(p); utils.append_log(product_id, change, reason, session.get('username')); flash('Stock ajustado')
    return redirect(url_for('index'))

@app.route('/create_order', methods=['GET','POST'])
@login_required
def create_order():
    products = utils.load_products()
    if request.method == 'POST':
        product_ids = request.form.getlist('product_id'); qtys = request.form.getlist('qty'); notes = request.form.getlist('note')
        items = []; total = 0.0; prod_map = {p['id']:p for p in products}
        for i,pid_raw in enumerate(product_ids):
            if not pid_raw: continue
            pid = int(pid_raw); qty = int(qtys[i] or 1); prod = prod_map.get(pid)
            if not prod: flash('Producto no encontrado'); return redirect(url_for('create_order'))
            if prod['stock'] < qty: flash(f'Stock insuficiente para {prod["name"]}'); return redirect(url_for('create_order'))
        for i,pid_raw in enumerate(product_ids):
            if not pid_raw: continue
            pid = int(pid_raw); qty = int(qtys[i] or 1); note = notes[i] if notes and i < len(notes) else ''
            prod = prod_map.get(pid); prod['stock'] -= qty; utils.save_product_row(prod); utils.append_log(pid, -qty, 'venta/pedido', session.get('username'))
            items.append({'id':pid,'name':prod['name'],'qty':qty,'price':prod['price'],'size':prod['size'],'color':prod['color'],'note':note}); total += prod['price'] * qty
        order = {'customer_name':request.form.get('customer_name'),'address':request.form.get('address'),'city':request.form.get('city'),'phone':request.form.get('phone'),'items':items,'total_price':total,'created_at':datetime.utcnow().isoformat()}
        utils.append_order(order); flash('Pedido registrado'); return redirect(url_for('orders'))
    return render_template('create_order.html', products=products)

@app.route('/orders')
@login_required
def orders():
    orders = utils.load_orders(200); return render_template('orders.html', orders=orders)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
