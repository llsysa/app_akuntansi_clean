import streamlit as st
import pandas as pd
from datetime import datetime
import json
import os
import re
import hashlib
from io import BytesIO
import uuid

# ===== FUNGSI to_excel =====
def to_excel(df):
    """Convert DataFrame to Excel file for download"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# ===== FUNGSI VALIDASI =====
def validate_email(email):
    pattern = r'^[a-zA-Z0-9._%+-]+@(gmail|yahoo)\.(com|co\.id)$'
    return re.match(pattern, email) is not None

def validate_password(password):
    if len(password) < 6:
        return False, "Password harus minimal 6 karakter!"
    return True, ""

# ===== CLASS USER MANAGER =====
class UserManager:
    def __init__(self):
        self.users_file = 'users.json'
        self.load_users()
    
    def load_users(self):
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'r', encoding='utf-8') as f:
                    self.users = json.load(f)
            except:
                self.users = {}
        else:
            self.users = {}
    
    def save_users(self):
        with open(self.users_file, 'w', encoding='utf-8') as f:
            json.dump(self.users, f, indent=4, ensure_ascii=False)
    
    def hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()
    
    def register_user(self, email, password):
        if email in self.users:
            return False, "Email sudah terdaftar!"
        
        self.users[email] = {
            'password_hash': self.hash_password(password),
            'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        self.save_users()
        return True, "Pendaftaran berhasil!"
    
    def verify_user(self, email, password):
        if email not in self.users:
            return False, "Email belum terdaftar!"
        
        if self.users[email]['password_hash'] == self.hash_password(password):
            return True, "Login berhasil!"
        else:
            return False, "Password salah!"

# ===== CLASS ACCOUNTING SYSTEM =====
class AccountingSystem:
    def __init__(self):
        self.accounts = {
            '101': {'name': 'Kas', 'type': 'Aset', 'balance': 0},
            '102': {'name': 'Bank', 'type': 'Aset', 'balance': 0},
            '115': {'name': 'Persediaan Caping', 'type': 'Aset', 'balance': 0},
            '116': {'name': 'Persediaan Cat', 'type': 'Aset', 'balance': 0},
            '117': {'name': 'Persediaan Kuas', 'type': 'Aset', 'balance': 0},
            '201': {'name': 'Utang Freelance', 'type': 'Kewajiban', 'balance': 0},
            '202': {'name': 'Utang Operasional', 'type': 'Kewajiban', 'balance': 0},
            '301': {'name': 'Ekuitas Pokdarwis', 'type': 'Modal', 'balance': 0},
            '401': {'name': 'Pendapatan Paket Wisata', 'type': 'Pendapatan', 'balance': 0},
            '402': {'name': 'Pendapatan Melukis Caping', 'type': 'Pendapatan', 'balance': 0},
            '403': {'name': 'Pendapatan Eksplorasi Singkong', 'type': 'Pendapatan', 'balance': 0},
            '404': {'name': 'Pendapatan Edukasi Pertanian', 'type': 'Pendapatan', 'balance': 0},
            '405': {'name': 'Pendapatan Konsumsi / Makan', 'type': 'Pendapatan', 'balance': 0},
            '406': {'name': 'Pendapatan Aktivitas Ikan & Sendang', 'type': 'Pendapatan', 'balance': 0},
            '407': {'name': 'Pendapatan Pemandu', 'type': 'Pendapatan', 'balance': 0},
            '501': {'name': 'Beban Melukis Caping', 'type': 'Beban', 'balance': 0},
            '502': {'name': 'Beban Snack / Break', 'type': 'Beban', 'balance': 0},
            '503': {'name': 'Beban Sendang', 'type': 'Beban', 'balance': 0},
            '504': {'name': 'Beban Pakan Kambing / Sapi', 'type': 'Beban', 'balance': 0},
            '505': {'name': 'Beban Kolam Ikan', 'type': 'Beban', 'balance': 0},
            '506': {'name': 'Beban Menanam Padi', 'type': 'Beban', 'balance': 0},
            '507': {'name': 'Beban Tangkap Ikan', 'type': 'Beban', 'balance': 0},
            '508': {'name': 'Beban Makan Siang', 'type': 'Beban', 'balance': 0},
            '509': {'name': 'Beban Pemandu Wisata', 'type': 'Beban', 'balance': 0},
            '510': {'name': 'Beban Pokdarwis', 'type': 'Beban', 'balance': 0},
            '511': {'name': 'Beban Marketing', 'type': 'Beban', 'balance': 0},
            '512': {'name': 'Beban Juru Kunci', 'type': 'Beban', 'balance': 0},
            '520': {'name': 'Beban Tenaga Kerja Freelance', 'type': 'Beban', 'balance': 0},
            '530': {'name': 'Beban Listrik', 'type': 'Beban', 'balance': 0},
            '531': {'name': 'Beban Internet', 'type': 'Beban', 'balance': 0},
            '532': {'name': 'Beban Kebersihan', 'type': 'Beban', 'balance': 0}
        }
        
        self.transactions = []
        self.adjustments = []
        self.load_data()
    
    def format_rupiah(self, amount):
        if amount == 0:
            return "Rp 0"
        return "Rp {:,.0f}".format(amount).replace(",", ".")
    
    def save_data(self):
        data = {
            'accounts': self.accounts,
            'transactions': self.transactions,
            'adjustments': self.adjustments
        }
        with open('accounting_data.json', 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    
    def load_data(self):
        if os.path.exists('accounting_data.json'):
            try:
                with open('accounting_data.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.accounts = data.get('accounts', self.accounts)
                    self.transactions = data.get('transactions', [])
                    self.adjustments = data.get('adjustments', [])
                    
                    # Pastikan semua transaksi dan penyesuaian memiliki ID
                    for trans in self.transactions:
                        if 'id' not in trans:
                            trans['id'] = str(uuid.uuid4())
                    
                    for adj in self.adjustments:
                        if 'id' not in adj:
                            adj['id'] = str(uuid.uuid4())
                            
            except Exception as e:
                st.warning(f"Error loading data: {e}. Starting with fresh data.")
    
    def add_transaction(self, date, desc, debit_acc, credit_acc, amount):
        if debit_acc not in self.accounts or credit_acc not in self.accounts:
            raise ValueError("Kode akun tidak valid")
        
        if amount <= 0:
            raise ValueError("Jumlah harus lebih dari 0")
        
        transaction = {
            'id': str(uuid.uuid4()),
            'date': date,
            'description': desc,
            'debit_account': debit_acc,
            'credit_account': credit_acc,
            'amount': amount
        }
        self.transactions.append(transaction)
        
        if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[debit_acc]['balance'] += amount
        else:
            self.accounts[debit_acc]['balance'] -= amount
        
        if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[credit_acc]['balance'] += amount
        else:
            self.accounts[credit_acc]['balance'] -= amount
        
        self.save_data()
    
    def update_transaction(self, index, date, desc, debit_acc, credit_acc, amount):
        if index < 0 or index >= len(self.transactions):
            raise ValueError("Index transaksi tidak valid")
        
        old_trans = self.transactions[index]
        self.reverse_transaction_effect(old_trans)
        
        self.transactions[index] = {
            'id': old_trans['id'],
            'date': date,
            'description': desc,
            'debit_account': debit_acc,
            'credit_account': credit_acc,
            'amount': amount
        }
        
        self.apply_transaction_effect(self.transactions[index])
        
        self.save_data()
    
    def reverse_transaction_effect(self, transaction):
        debit_acc = transaction['debit_account']
        credit_acc = transaction['credit_account']
        amount = transaction['amount']
        
        if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[debit_acc]['balance'] -= amount
        else:
            self.accounts[debit_acc]['balance'] += amount
        
        if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[credit_acc]['balance'] -= amount
        else:
            self.accounts[credit_acc]['balance'] += amount
    
    def apply_transaction_effect(self, transaction):
        debit_acc = transaction['debit_account']
        credit_acc = transaction['credit_account']
        amount = transaction['amount']
        
        if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[debit_acc]['balance'] += amount
        else:
            self.accounts[debit_acc]['balance'] -= amount
        
        if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[credit_acc]['balance'] += amount
        else:
            self.accounts[credit_acc]['balance'] -= amount
    
    def delete_transaction_by_id(self, transaction_id):
        for i, trans in enumerate(self.transactions):
            if trans['id'] == transaction_id:
                self.reverse_transaction_effect(trans)
                self.transactions.pop(i)
                self.save_data()
                return True
        return False
    
    def add_adjustment(self, date, desc, debit_acc, credit_acc, amount):
        if debit_acc not in self.accounts or credit_acc not in self.accounts:
            raise ValueError("Kode akun tidak valid")
        
        if amount <= 0:
            raise ValueError("Jumlah harus lebih dari 0")
        
        adjustment = {
            'id': str(uuid.uuid4()),
            'date': date,
            'description': desc,
            'debit_account': debit_acc,
            'credit_account': credit_acc,
            'amount': amount
        }
        self.adjustments.append(adjustment)
        
        # Apply adjustment effect to accounts
        if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[debit_acc]['balance'] += amount
        else:
            self.accounts[debit_acc]['balance'] -= amount
        
        if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[credit_acc]['balance'] += amount
        else:
            self.accounts[credit_acc]['balance'] -= amount
        
        self.save_data()
    
    def update_adjustment(self, index, date, desc, debit_acc, credit_acc, amount):
        if index < 0 or index >= len(self.adjustments):
            raise ValueError("Index penyesuaian tidak valid")
        
        # Simpan data lama untuk reverse effect
        old_adj = self.adjustments[index]
        
        # Reverse efek penyesuaian lama
        old_debit_acc = old_adj['debit_account']
        old_credit_acc = old_adj['credit_account']
        old_amount = old_adj['amount']
        
        if self.accounts[old_debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[old_debit_acc]['balance'] -= old_amount
        else:
            self.accounts[old_debit_acc]['balance'] += old_amount
        
        if self.accounts[old_credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[old_credit_acc]['balance'] -= old_amount
        else:
            self.accounts[old_credit_acc]['balance'] += old_amount
        
        # Update data
        self.adjustments[index] = {
            'id': old_adj['id'],
            'date': date,
            'description': desc,
            'debit_account': debit_acc,
            'credit_account': credit_acc,
            'amount': amount
        }
        
        # Apply efek penyesuaian baru
        if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
            self.accounts[debit_acc]['balance'] += amount
        else:
            self.accounts[debit_acc]['balance'] -= amount
        
        if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
            self.accounts[credit_acc]['balance'] += amount
        else:
            self.accounts[credit_acc]['balance'] -= amount
        
        self.save_data()
    
    def delete_adjustment_by_id(self, adjustment_id):
        """Menghapus penyesuaian berdasarkan ID"""
        for i, adj in enumerate(self.adjustments):
            if adj['id'] == adjustment_id:
                # Reverse the effect of adjustment on accounts
                debit_acc = adj['debit_account']
                credit_acc = adj['credit_account']
                amount = adj['amount']
                
                if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
                    self.accounts[debit_acc]['balance'] -= amount
                else:
                    self.accounts[debit_acc]['balance'] += amount
                
                if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
                    self.accounts[credit_acc]['balance'] -= amount
                else:
                    self.accounts[credit_acc]['balance'] += amount
                
                self.adjustments.pop(i)
                self.save_data()
                return True
        return False

    def delete_all_adjustments(self):
        """Menghapus semua penyesuaian sekaligus"""
        if not self.adjustments:
            return 0
        
        total_deleted = len(self.adjustments)
        
        # Reverse effect dari semua penyesuaian
        for adj in self.adjustments:
            debit_acc = adj['debit_account']
            credit_acc = adj['credit_account']
            amount = adj['amount']
            
            if self.accounts[debit_acc]['type'] in ['Aset', 'Beban']:
                self.accounts[debit_acc]['balance'] -= amount
            else:
                self.accounts[debit_acc]['balance'] += amount
            
            if self.accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
                self.accounts[credit_acc]['balance'] -= amount
            else:
                self.accounts[credit_acc]['balance'] += amount
        
        # Kosongkan semua penyesuaian
        self.adjustments.clear()
        self.save_data()
        
        return total_deleted

    def get_adjusted_accounts(self):
        """Mengembalikan saldo akun setelah penyesuaian dengan efek yang benar"""
        # Mulai dari saldo normal terlebih dahulu
        adjusted_accounts = {code: info.copy() for code, info in self.accounts.items()}
        
        # Terapkan efek penyesuaian
        for adj in self.adjustments:
            debit_acc = adj['debit_account']
            credit_acc = adj['credit_account']
            amount = adj['amount']
            
            # Untuk akun debit (Aset/Beban) - bertambah di debit
            if adjusted_accounts[debit_acc]['type'] in ['Aset', 'Beban']:
                adjusted_accounts[debit_acc]['balance'] += amount
            else:
                adjusted_accounts[debit_acc]['balance'] -= amount
            
            # Untuk akun kredit (Kewajiban/Modal/Pendapatan) - bertambah di kredit
            if adjusted_accounts[credit_acc]['type'] in ['Kewajiban', 'Modal', 'Pendapatan']:
                adjusted_accounts[credit_acc]['balance'] += amount
            else:
                adjusted_accounts[credit_acc]['balance'] -= amount
        
        return adjusted_accounts

    def get_adjusting_journal(self):
        """Menampilkan jurnal penyesuaian"""
        journal = []
        for i, adj in enumerate(self.adjustments):
            journal.append({
                'No': i + 1,
                'Tanggal': adj['date'],
                'Keterangan': adj['description'],
                'Nama Akun': self.accounts[adj['debit_account']]['name'],
                'Ref': adj['debit_account'],
                'Debit': self.format_rupiah(adj['amount']),
                'Kredit': self.format_rupiah(0)
            })
            journal.append({
                'No': '',
                'Tanggal': '',
                'Keterangan': '',
                'Nama Akun': self.accounts[adj['credit_account']]['name'],
                'Ref': adj['credit_account'],
                'Debit': self.format_rupiah(0),
                'Kredit': self.format_rupiah(adj['amount'])
            })
        return pd.DataFrame(journal)
    
    def get_general_journal(self):
        journal = []
        for i, trans in enumerate(self.transactions):
            journal.append({
                'No': i + 1,
                'Tanggal': trans['date'],
                'Keterangan': trans['description'],
                'Nama Akun': self.accounts[trans['debit_account']]['name'],
                'Ref': trans['debit_account'],
                'Debit': self.format_rupiah(trans['amount']),
                'Kredit': self.format_rupiah(0)
            })
            journal.append({
                'No': '',
                'Tanggal': '',
                'Keterangan': '',
                'Nama Akun': self.accounts[trans['credit_account']]['name'],
                'Ref': trans['credit_account'],
                'Debit': self.format_rupiah(0),
                'Kredit': self.format_rupiah(trans['amount'])
            })
        return pd.DataFrame(journal)
    
    def get_ledger(self):
        ledger_data = []
        for acc_code, acc_info in sorted(self.accounts.items()):
            if acc_info['balance'] != 0:
                ledger_data.append({
                    'Kode Akun': acc_code,
                    'Nama Akun': acc_info['name'],
                    'Tipe': acc_info['type'],
                    'Saldo': self.format_rupiah(acc_info['balance'])
                })
        return pd.DataFrame(ledger_data)
    
    def get_detailed_ledger(self):
        detailed_ledger = {}
        
        for acc_code, acc_info in self.accounts.items():
            detailed_ledger[acc_code] = {
                'name': acc_info['name'],
                'type': acc_info['type'],
                'transactions': [],
                'balance': 0
            }
        
        for trans in self.transactions:
            detailed_ledger[trans['debit_account']]['transactions'].append({
                'date': trans['date'],
                'description': trans['description'],
                'ref': trans['credit_account'],
                'debit': trans['amount'],
                'credit': 0
            })
            detailed_ledger[trans['credit_account']]['transactions'].append({
                'date': trans['date'],
                'description': trans['description'],
                'ref': trans['debit_account'],
                'debit': 0,
                'credit': trans['amount']
            })
        
        for acc_code in detailed_ledger:
            balance = 0
            for trans in detailed_ledger[acc_code]['transactions']:
                if detailed_ledger[acc_code]['type'] in ['Aset', 'Beban']:
                    balance += trans['debit'] - trans['credit']
                else:
                    balance += trans['credit'] - trans['debit']
                trans['balance'] = balance
            detailed_ledger[acc_code]['balance'] = balance
        
        return detailed_ledger
    
    def get_trial_balance(self):
        trial_data = []
        total_debit = 0
        total_credit = 0
        
        for acc_code, acc_info in sorted(self.accounts.items()):
            balance = acc_info['balance']
            
            if acc_info['type'] in ['Aset', 'Beban']:
                if balance >= 0:
                    debit_amount = balance
                    credit_amount = 0
                else:
                    debit_amount = 0
                    credit_amount = abs(balance)
            else:
                if balance >= 0:
                    debit_amount = 0
                    credit_amount = balance
                else:
                    debit_amount = abs(balance)
                    credit_amount = 0
            
            if debit_amount > 0 or credit_amount > 0:
                trial_data.append({
                    'No Akun': acc_code,
                    'Nama Akun': acc_info['name'],
                    'Debit': self.format_rupiah(debit_amount),
                    'Kredit': self.format_rupiah(credit_amount)
                })
                total_debit += debit_amount
                total_credit += credit_amount
        
        df = pd.DataFrame(trial_data)
        return df, total_debit, total_credit
    
    def get_adjusted_trial_balance(self):
        adjusted_accounts = self.get_adjusted_accounts()
        
        trial_data = []
        total_debit = 0
        total_credit = 0
        
        for acc_code, acc_info in sorted(adjusted_accounts.items()):
            balance = acc_info['balance']
            
            if acc_info['type'] in ['Aset', 'Beban']:
                if balance >= 0:
                    debit_amount = balance
                    credit_amount = 0
                else:
                    debit_amount = 0
                    credit_amount = abs(balance)
            else:
                if balance >= 0:
                    debit_amount = 0
                    credit_amount = balance
                else:
                    debit_amount = abs(balance)
                    credit_amount = 0
            
            if debit_amount > 0 or credit_amount > 0:
                trial_data.append({
                    'No Akun': acc_code,
                    'Nama Akun': acc_info['name'],
                    'Debit': self.format_rupiah(debit_amount),
                    'Kredit': self.format_rupiah(credit_amount)
                })
                total_debit += debit_amount
                total_credit += credit_amount
        
        df = pd.DataFrame(trial_data)
        return df, total_debit, total_credit
    
    def get_financial_statements(self):
        """Laporan keuangan menggunakan data setelah penyesuaian"""
        adjusted_accounts = self.get_adjusted_accounts()
        
        total_revenue = 0
        total_expense = 0
        
        # Hitung pendapatan dan beban dari akun yang sudah disesuaikan
        for acc_code, acc_info in adjusted_accounts.items():
            if acc_info['type'] == 'Pendapatan':
                total_revenue += acc_info['balance']
            elif acc_info['type'] == 'Beban':
                total_expense += acc_info['balance']
        
        net_income = total_revenue - total_expense
        
        # Siapkan data untuk neraca
        assets_data = []
        liabilities_data = []
        equity_data = []
        
        total_assets = 0
        total_liabilities = 0
        total_equity = 0
        
        for acc_code, acc_info in sorted(adjusted_accounts.items()):
            balance = acc_info['balance']
            
            if acc_info['type'] == 'Aset':
                if balance > 0:
                    assets_data.append({
                        'Akun': f"{acc_code} - {acc_info['name']}",
                        'Jumlah': self.format_rupiah(balance)
                    })
                    total_assets += balance
                elif balance < 0:
                    # Jika saldo aset negatif, tampilkan sebagai koreksi
                    assets_data.append({
                        'Akun': f"{acc_code} - {acc_info['name']} (Koreksi)",
                        'Jumlah': self.format_rupiah(balance)
                    })
                    total_assets += balance
            
            elif acc_info['type'] == 'Kewajiban':
                if balance > 0:
                    liabilities_data.append({
                        'Akun': f"{acc_code} - {acc_info['name']}",
                        'Jumlah': self.format_rupiah(balance)
                    })
                    total_liabilities += balance
                elif balance < 0:
                    liabilities_data.append({
                        'Akun': f"{acc_code} - {acc_info['name']} (Koreksi)",
                        'Jumlah': self.format_rupiah(balance)
                    })
                    total_liabilities += balance
            
            elif acc_info['type'] == 'Modal':
                # Untuk modal, tambahkan laba/rugi bersih
                modal_balance = balance + net_income
                if modal_balance != 0:
                    equity_data.append({
                        'Akun': f"{acc_code} - {acc_info['name']}",
                        'Jumlah': self.format_rupiah(modal_balance)
                    })
                    total_equity += modal_balance
        
        return {
            'income_statement': {
                'revenue': total_revenue,
                'expense': total_expense,
                'net_income': net_income
            },
            'balance_sheet': {
                'assets': pd.DataFrame(assets_data),
                'liabilities': pd.DataFrame(liabilities_data),
                'equity': pd.DataFrame(equity_data),
                'total_assets': total_assets,
                'total_liabilities': total_liabilities,
                'total_equity': total_equity,
                'total_liabilities_equity': total_liabilities + total_equity
            }
        }

# ===== FUNGSI TAMPILAN MENU =====
def show_coa_page(accounting_system):
    st.header("📋 Chart of Accounts (COA)")
    
    coa_data = []
    for acc_code, acc_info in sorted(accounting_system.accounts.items()):
        coa_data.append({
            'Kode Akun': acc_code,
            'Nama Akun': acc_info['name'],
            'Tipe': acc_info['type'],
            'Saldo': accounting_system.format_rupiah(acc_info['balance'])
        })
    
    coa_df = pd.DataFrame(coa_data)
    st.dataframe(coa_df, use_container_width=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        total_akun = len(accounting_system.accounts)
        st.metric("Total Akun", total_akun)
    with col2:
        akun_aktif = len([acc for acc in accounting_system.accounts.values() if acc['balance'] != 0])
        st.metric("Akun Aktif", akun_aktif)
    with col3:
        total_debit = sum(acc['balance'] for acc in accounting_system.accounts.values() if acc['balance'] > 0)
        st.metric("Total Debit", accounting_system.format_rupiah(total_debit))
    with col4:
        total_kredit = abs(sum(acc['balance'] for acc in accounting_system.accounts.values() if acc['balance'] < 0))
        st.metric("Total Kredit", accounting_system.format_rupiah(total_kredit))

def show_input_transaksi_page(accounting_system):
    st.header("📥 Input Transaksi")
    
    # Inisialisasi session state untuk mencegah duplikasi
    if 'form_submitted' not in st.session_state:
        st.session_state.form_submitted = False
    if 'last_transaction_id' not in st.session_state:
        st.session_state.last_transaction_id = None
    
    tab1, tab2, tab3 = st.tabs(["➕ Input Baru", "✏️ Edit Transaksi", "🗑️ Hapus Transaksi"])
    
    with tab1:
        with st.form("transaction_form", clear_on_submit=True):  # clear_on_submit untuk mencegah duplikasi
            col1, col2 = st.columns(2)
            
            with col1:
                trans_date = st.date_input("Tanggal Transaksi", datetime.now())
                trans_desc = st.text_input("Keterangan Transaksi", placeholder="Contoh: Pembayaran paket wisata")
                amount = st.number_input("Jumlah (Rp)", min_value=0, step=1000, format="%d")
            
            with col2:
                account_options = [f"{code} - {info['name']}" for code, info in accounting_system.accounts.items()]
                debit_acc = st.selectbox("Akun Debit", account_options, key="debit")
                credit_acc = st.selectbox("Akun Kredit", account_options, key="credit")
                
                st.info("**Pastikan:**")
                st.info("- Akun debit dan kredit berbeda")
                st.info("- Jumlah lebih dari 0")
            
            submitted = st.form_submit_button("💾 Simpan Transaksi", use_container_width=True)
            
            if submitted:
                # Cek apakah form sudah disubmit sebelumnya
                current_transaction_key = f"{trans_date}_{trans_desc}_{amount}_{debit_acc}_{credit_acc}"
                
                if (st.session_state.form_submitted and 
                    st.session_state.last_transaction_id == current_transaction_key):
                    st.warning("⏳ Transaksi sedang diproses, harap tunggu...")
                    return
                
                if not trans_desc:
                    st.error("❌ Keterangan transaksi harus diisi!")
                elif amount <= 0:
                    st.error("❌ Jumlah harus lebih dari 0!")
                elif debit_acc == credit_acc:
                    st.error("❌ Akun debit dan kredit tidak boleh sama!")
                else:
                    try:
                        # Set status form submitted
                        st.session_state.form_submitted = True
                        st.session_state.last_transaction_id = current_transaction_key
                        
                        debit_code = debit_acc.split(" - ")[0]
                        credit_code = credit_acc.split(" - ")[0]
                        
                        accounting_system.add_transaction(
                            trans_date.strftime("%Y-%m-%d"),
                            trans_desc,
                            debit_code,
                            credit_code,
                            amount
                        )
                        st.success("✅ Transaksi berhasil disimpan!")
                        st.balloons()
                        
                        # Reset form setelah berhasil disimpan
                        st.session_state.form_submitted = False
                        st.session_state.last_transaction_id = None
                        
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan transaksi: {str(e)}")
                        # Reset status jika gagal
                        st.session_state.form_submitted = False
    
    with tab2:
        st.subheader("Edit Transaksi")
        
        if not accounting_system.transactions:
            st.info("📝 Belum ada transaksi yang bisa diedit")
        else:
            trans_options = []
            for i, trans in enumerate(accounting_system.transactions):
                trans_options.append(f"{i+1}. {trans['date']} - {trans['description']} - {accounting_system.format_rupiah(trans['amount'])}")
            
            selected_trans = st.selectbox("Pilih Transaksi untuk Edit:", trans_options, key="edit_select")
            
            if selected_trans:
                trans_index = trans_options.index(selected_trans)
                trans = accounting_system.transactions[trans_index]
                
                with st.form("edit_transaction_form"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        edit_date = st.date_input("Tanggal Transaksi", datetime.strptime(trans['date'], "%Y-%m-%d"))
                        edit_desc = st.text_input("Keterangan Transaksi", value=trans['description'])
                        edit_amount = st.number_input("Jumlah (Rp)", min_value=0, step=1000, format="%d", value=trans['amount'])
                    
                    with col2:
                        account_options = [f"{code} - {info['name']}" for code, info in accounting_system.accounts.items()]
                        
                        current_debit = f"{trans['debit_account']} - {accounting_system.accounts[trans['debit_account']]['name']}"
                        current_credit = f"{trans['credit_account']} - {accounting_system.accounts[trans['credit_account']]['name']}"
                        
                        edit_debit = st.selectbox("Akun Debit", account_options, 
                                                index=account_options.index(current_debit) if current_debit in account_options else 0,
                                                key="edit_debit")
                        edit_credit = st.selectbox("Akun Kredit", account_options,
                                                 index=account_options.index(current_credit) if current_credit in account_options else 0,
                                                 key="edit_credit")
                    
                    update_submitted = st.form_submit_button("💾 Update Transaksi", use_container_width=True)
                    
                    if update_submitted:
                        if not edit_desc:
                            st.error("❌ Keterangan transaksi harus diisi!")
                        elif edit_amount <= 0:
                            st.error("❌ Jumlah harus lebih dari 0!")
                        elif edit_debit == edit_credit:
                            st.error("❌ Akun debit dan kredit tidak boleh sama!")
                        else:
                            try:
                                debit_code = edit_debit.split(" - ")[0]
                                credit_code = edit_credit.split(" - ")[0]
                                
                                accounting_system.update_transaction(
                                    trans_index,
                                    edit_date.strftime("%Y-%m-%d"),
                                    edit_desc,
                                    debit_code,
                                    credit_code,
                                    edit_amount
                                )
                                st.success("✅ Transaksi berhasil diupdate!")
                                st.balloons()
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"❌ Gagal mengupdate transaksi: {str(e)}")
    
    with tab3:
        st.subheader("Hapus Transaksi")
        
        if not accounting_system.transactions:
            st.info("📝 Belum ada transaksi yang bisa dihapus")
        else:
            st.warning("⚠️ Peringatan: Menghapus transaksi akan mempengaruhi semua laporan!")
            
            trans_to_delete = []
            for i, trans in enumerate(accounting_system.transactions):
                col1, col2, col3 = st.columns([1, 4, 2])
                with col1:
                    # Gunakan index sebagai fallback jika id tidak ada
                    delete_key = f"delete_{trans.get('id', i)}"
                    delete = st.checkbox("", key=delete_key)
                with col2:
                    st.write(f"**{trans['date']}** - {trans['description']}")
                    st.write(f"Debit: {trans['debit_account']} | Kredit: {trans['credit_account']}")
                with col3:
                    st.write(f"**{accounting_system.format_rupiah(trans['amount'])}**")
                
                if delete:
                    trans_to_delete.append(trans['id'])
            
            if trans_to_delete:
                if st.button("🗑️ Hapus Transaksi Terpilih", type="secondary", use_container_width=True):
                    success_count = 0
                    for trans_id in trans_to_delete:
                        if accounting_system.delete_transaction_by_id(trans_id):
                            success_count += 1
                    
                    st.success(f"✅ {success_count} transaksi berhasil dihapus!")
                    st.rerun()

def show_jurnal_umum_page(accounting_system):
    st.header("📔 Jurnal Umum")
    
    journal_df = accounting_system.get_general_journal()
    if not journal_df.empty:
        st.dataframe(journal_df, use_container_width=True)
        
        csv = journal_df.to_csv(index=False, encoding='utf-8')
        st.download_button(
            label="📥 Download Jurnal Umum (CSV)",
            data=csv,
            file_name="jurnal_umum.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        col1, col2 = st.columns(2)
        with col1:
            total_entries = len(journal_df)
            st.metric("Total Entri Jurnal", total_entries)
        with col2:
            total_transactions = len(accounting_system.transactions)
            st.metric("Total Transaksi", total_transactions)
    else:
        st.info("📝 Belum ada transaksi yang dicatat")

def show_buku_besar_page(accounting_system):
    st.header("📒 Buku Besar")
    
    tab1, tab2 = st.tabs(["📋 Ringkasan Saldo", "📖 Buku Besar Detail"])
    
    with tab1:
        ledger_df = accounting_system.get_ledger()
        if not ledger_df.empty:
            st.dataframe(ledger_df, use_container_width=True)
        else:
            st.info("📝 Belum ada saldo akun")
    
    with tab2:
        detailed_ledger = accounting_system.get_detailed_ledger()
        accounts_with_balance = [code for code, info in detailed_ledger.items() if info['balance'] != 0]
        
        if accounts_with_balance:
            selected_account = st.selectbox(
                "Pilih Akun:",
                accounts_with_balance,
                format_func=lambda x: f"{x} - {detailed_ledger[x]['name']} (Saldo: {accounting_system.format_rupiah(detailed_ledger[x]['balance'])})"
            )
            
            if selected_account:
                account_data = detailed_ledger[selected_account]
                transactions_df = pd.DataFrame(account_data['transactions'])
                
                if not transactions_df.empty:
                    transactions_df['Debit'] = transactions_df['debit'].apply(accounting_system.format_rupiah)
                    transactions_df['Kredit'] = transactions_df['credit'].apply(accounting_system.format_rupiah)
                    transactions_df['Saldo'] = transactions_df['balance'].apply(accounting_system.format_rupiah)
                    
                    display_df = transactions_df[['date', 'description', 'ref', 'Debit', 'Kredit', 'Saldo']]
                    display_df.columns = ['Tanggal', 'Keterangan', 'Ref', 'Debit', 'Kredit', 'Saldo']
                    
                    st.dataframe(display_df, use_container_width=True)
                else:
                    st.info("📝 Tidak ada transaksi untuk akun ini")
        else:
            st.info("📝 Belum ada akun dengan saldo")

def show_neraca_saldo_page(accounting_system):
    st.header("⚖️ Neraca Saldo")
    
    trial_df, total_debit, total_credit = accounting_system.get_trial_balance()
    
    if not trial_df.empty:
        st.dataframe(trial_df, use_container_width=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Debit", accounting_system.format_rupiah(total_debit))
        with col2:
            st.metric("Total Kredit", accounting_system.format_rupiah(total_credit))
        with col3:
            balance_status = "✅ Balance" if total_debit == total_credit else "❌ Tidak Balance"
            st.metric("Status", balance_status)
        
        if total_debit != total_credit:
            st.error(f"⚠️ Selisih: {accounting_system.format_rupiah(abs(total_debit - total_credit))}")
    else:
        st.info("📝 Belum ada saldo akun")

def show_penyesuaian_page(accounting_system):
    st.header("🔄 Jurnal Penyesuaian")
    
    # Inisialisasi session state untuk mencegah duplikasi penyesuaian
    if 'adjustment_form_submitted' not in st.session_state:
        st.session_state.adjustment_form_submitted = False
    if 'last_adjustment_id' not in st.session_state:
        st.session_state.last_adjustment_id = None
    
    tab1, tab2, tab3, tab4 = st.tabs(["➕ Input Penyesuaian", "✏️ Edit Penyesuaian", "🗑️ Hapus Penyesuaian", "📋 Lihat Jurnal Penyesuaian"])
    
    with tab1:
        with st.form("adjustment_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                adj_date = st.date_input("Tanggal Penyesuaian", datetime.now(), key="adj_date")
                adj_desc = st.text_input("Keterangan Penyesuaian", placeholder="Contoh: Penyusutan peralatan")
                adj_amount = st.number_input("Jumlah Penyesuaian (Rp)", min_value=0, step=1000, format="%d", key="adj_amount")
            
            with col2:
                account_options = [f"{code} - {info['name']}" for code, info in accounting_system.accounts.items()]
                adj_debit = st.selectbox("Akun Debit Penyesuaian", account_options, key="adj_debit")
                adj_credit = st.selectbox("Akun Kredit Penyesuaian", account_options, key="adj_credit")
                
                st.info("**Pastikan:**")
                st.info("- Akun debit dan kredit berbeda")
                st.info("- Jumlah lebih dari 0")
            
            submitted = st.form_submit_button("💾 Simpan Penyesuaian", use_container_width=True)
            
            if submitted:
                # Cek apakah form sudah disubmit sebelumnya
                current_adjustment_key = f"{adj_date}_{adj_desc}_{adj_amount}_{adj_debit}_{adj_credit}"
                
                if (st.session_state.adjustment_form_submitted and 
                    st.session_state.last_adjustment_id == current_adjustment_key):
                    st.warning("⏳ Penyesuaian sedang diproses, harap tunggu...")
                    return
                
                if not adj_desc:
                    st.error("❌ Keterangan penyesuaian harus diisi!")
                elif adj_amount <= 0:
                    st.error("❌ Jumlah harus lebih dari 0!")
                elif adj_debit == adj_credit:
                    st.error("❌ Akun debit dan kredit tidak boleh sama!")
                else:
                    try:
                        # Set status form submitted
                        st.session_state.adjustment_form_submitted = True
                        st.session_state.last_adjustment_id = current_adjustment_key
                        
                        debit_code = adj_debit.split(" - ")[0]
                        credit_code = adj_credit.split(" - ")[0]
                        
                        accounting_system.add_adjustment(
                            adj_date.strftime("%Y-%m-%d"),
                            adj_desc,
                            debit_code,
                            credit_code,
                            adj_amount
                        )
                        st.success("✅ Penyesuaian berhasil disimpan!")
                        st.balloons()
                        
                        # Reset form setelah berhasil disimpan
                        st.session_state.adjustment_form_submitted = False
                        st.session_state.last_adjustment_id = None
                        
                    except Exception as e:
                        st.error(f"❌ Gagal menyimpan penyesuaian: {str(e)}")
                        # Reset status jika gagal
                        st.session_state.adjustment_form_submitted = False
    
    with tab2:
        st.subheader("Edit Penyesuaian")
        
        if not accounting_system.adjustments:
            st.info("📝 Belum ada penyesuaian yang bisa diedit")
        else:
            adj_options = []
            for i, adj in enumerate(accounting_system.adjustments):
                adj_options.append(f"{i+1}. {adj['date']} - {adj['description']} - {accounting_system.format_rupiah(adj['amount'])}")
            
            selected_adj = st.selectbox("Pilih Penyesuaian untuk Edit:", adj_options, key="edit_adj_select")
            
            if selected_adj:
                adj_index = adj_options.index(selected_adj)
                adj = accounting_system.adjustments[adj_index]
                
                with st.form("edit_adjustment_form"):
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        edit_adj_date = st.date_input("Tanggal Penyesuaian", 
                                                    datetime.strptime(adj['date'], "%Y-%m-%d"),
                                                    key="edit_adj_date")
                        edit_adj_desc = st.text_input("Keterangan Penyesuaian", 
                                                    value=adj['description'],
                                                    key="edit_adj_desc")
                        edit_adj_amount = st.number_input("Jumlah Penyesuaian (Rp)", 
                                                        min_value=0, step=1000, format="%d", 
                                                        value=adj['amount'],
                                                        key="edit_adj_amount")
                    
                    with col2:
                        account_options = [f"{code} - {info['name']}" for code, info in accounting_system.accounts.items()]
                        
                        current_debit = f"{adj['debit_account']} - {accounting_system.accounts[adj['debit_account']]['name']}"
                        current_credit = f"{adj['credit_account']} - {accounting_system.accounts[adj['credit_account']]['name']}"
                        
                        edit_adj_debit = st.selectbox("Akun Debit Penyesuaian", account_options,
                                                    index=account_options.index(current_debit) if current_debit in account_options else 0,
                                                    key="edit_adj_debit")
                        edit_adj_credit = st.selectbox("Akun Kredit Penyesuaian", account_options,
                                                     index=account_options.index(current_credit) if current_credit in account_options else 0,
                                                     key="edit_adj_credit")
                    
                    update_submitted = st.form_submit_button("💾 Update Penyesuaian", use_container_width=True)
                    
                    if update_submitted:
                        if not edit_adj_desc:
                            st.error("❌ Keterangan penyesuaian harus diisi!")
                        elif edit_adj_amount <= 0:
                            st.error("❌ Jumlah harus lebih dari 0!")
                        elif edit_adj_debit == edit_adj_credit:
                            st.error("❌ Akun debit dan kredit tidak boleh sama!")
                        else:
                            try:
                                debit_code = edit_adj_debit.split(" - ")[0]
                                credit_code = edit_adj_credit.split(" - ")[0]
                                
                                accounting_system.update_adjustment(
                                    adj_index,
                                    edit_adj_date.strftime("%Y-%m-%d"),
                                    edit_adj_desc,
                                    debit_code,
                                    credit_code,
                                    edit_adj_amount
                                )
                                st.success("✅ Penyesuaian berhasil diupdate!")
                                st.balloons()
                                st.rerun()
                                
                            except Exception as e:
                                st.error(f"❌ Gagal mengupdate penyesuaian: {str(e)}")
    
    with tab3:
        st.subheader("Hapus Penyesuaian")
        
        if not accounting_system.adjustments:
            st.info("📝 Belum ada penyesuaian yang bisa dihapus")
        else:
            st.warning("⚠️ Peringatan: Menghapus penyesuaian akan mempengaruhi neraca saldo dan laporan keuangan!")
            
            # Tampilkan statistik penyesuaian
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Penyesuaian", len(accounting_system.adjustments))
            with col2:
                total_amount = sum(adj['amount'] for adj in accounting_system.adjustments)
                st.metric("Total Jumlah Penyesuaian", accounting_system.format_rupiah(total_amount))
            
            st.markdown("---")
            
            # Opsi hapus per penyesuaian
            st.subheader("Hapus Per Penyesuaian")
            adj_to_delete = []
            for i, adj in enumerate(accounting_system.adjustments):
                col1, col2, col3 = st.columns([1, 4, 2])
                with col1:
                    # PERBAIKAN: Gunakan get() untuk menghindari KeyError dan gunakan index sebagai fallback
                    adj_id = adj.get('id', f"adj_{i}")  # Jika tidak ada id, gunakan index sebagai fallback
                    delete = st.checkbox("", key=f"delete_adj_{adj_id}")
                with col2:
                    st.write(f"**{adj['date']}** - {adj['description']}")
                    st.write(f"Debit: {adj['debit_account']} | Kredit: {adj['credit_account']}")
                with col3:
                    st.write(f"**{accounting_system.format_rupiah(adj['amount'])}**")
                
                if delete:
                    # Pastikan kita menggunakan id yang benar
                    if 'id' in adj:
                        adj_to_delete.append(adj['id'])
                    else:
                        # Jika tidak ada id, coba hapus berdasarkan index
                        st.warning(f"Penyesuaian {i+1} tidak memiliki ID yang valid")
            
            if adj_to_delete:
                if st.button("🗑️ Hapus Penyesuaian Terpilih", type="secondary", use_container_width=True):
                    success_count = 0
                    for adj_id in adj_to_delete:
                        if accounting_system.delete_adjustment_by_id(adj_id):
                            success_count += 1
                    
                    st.success(f"✅ {success_count} penyesuaian berhasil dihapus!")
                    st.rerun()
            
            st.markdown("---")
            
            # Opsi hapus semua penyesuaian
            st.subheader("Hapus Semua Penyesuaian")
            st.error("🚨 **PERINGATAN TINGGI**: Tindakan ini akan menghapus SEMUA penyesuaian sekaligus dan tidak dapat dibatalkan!")
            
            col1, col2 = st.columns(2)
            with col1:
                # Konfirmasi dengan checkbox
                confirm_delete_all = st.checkbox("Saya mengerti dan ingin menghapus SEMUA penyesuaian")
            with col2:
                # Input konfirmasi teks
                text_confirm = st.text_input("Ketik 'HAPUS SEMUA' untuk konfirmasi:", placeholder="HAPUS SEMUA")
            
            if confirm_delete_all and text_confirm == "HAPUS SEMUA":
                if st.button("💥 HAPUS SEMUA PENYESUAIAN", type="primary", use_container_width=True):
                    total_deleted = accounting_system.delete_all_adjustments()
                    if total_deleted > 0:
                        st.success(f"✅ Semua {total_deleted} penyesuaian berhasil dihapus!")
                        st.balloons()
                        st.rerun()
                    else:
                        st.info("📝 Tidak ada penyesuaian yang dihapus")
            elif confirm_delete_all and text_confirm != "HAPUS SEMUA":
                st.warning("❌ Silakan ketik 'HAPUS SEMUA' untuk mengonfirmasi penghapusan")
    
    with tab4:
        st.subheader("Jurnal Penyesuaian")
        
        if not accounting_system.adjustments:
            st.info("📝 Belum ada jurnal penyesuaian")
        else:
            adjusting_journal_df = accounting_system.get_adjusting_journal()
            st.dataframe(adjusting_journal_df, use_container_width=True)
            
            col1, col2 = st.columns(2)
            with col1:
                total_adjustments = len(accounting_system.adjustments)
                st.metric("Total Penyesuaian", total_adjustments)
            with col2:
                total_amount = sum(adj['amount'] for adj in accounting_system.adjustments)
                st.metric("Total Jumlah Penyesuaian", accounting_system.format_rupiah(total_amount))

def show_neraca_saldo_penyesuaian_page(accounting_system):
    st.header("⚖️ Neraca Saldo Setelah Penyesuaian")
    
    adj_trial_df, total_debit, total_credit = accounting_system.get_adjusted_trial_balance()
    
    if not adj_trial_df.empty:
        st.dataframe(adj_trial_df, use_container_width=True)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Debit", accounting_system.format_rupiah(total_debit))
        with col2:
            st.metric("Total Kredit", accounting_system.format_rupiah(total_credit))
        with col3:
            balance_status = "✅ Balance" if total_debit == total_credit else "❌ Tidak Balance"
            st.metric("Status", balance_status)
        
        if total_debit != total_credit:
            st.error(f"⚠️ Selisih: {accounting_system.format_rupiah(abs(total_debit - total_credit))}")
    else:
        st.info("📝 Belum ada saldo akun setelah penyesuaian")

def show_laporan_keuangan_page(accounting_system):
    st.header("📊 Laporan Keuangan")
    
    financials = accounting_system.get_financial_statements()
    
    tab1, tab2, tab3 = st.tabs(["📈 Laporan Laba Rugi", "🏦 Neraca", "📥 Download Excel"])
    
    with tab1:
        st.subheader("Laporan Laba Rugi")
        st.markdown("---")
        
        revenue = financials['income_statement']['revenue']
        expense = financials['income_statement']['expense']
        net_income = financials['income_statement']['net_income']
        
        st.metric("Total Pendapatan", accounting_system.format_rupiah(revenue))
        st.metric("Total Beban", accounting_system.format_rupiah(expense))
        st.markdown("---")
        
        st.metric(
            "Laba/Rugi Bersih", 
            accounting_system.format_rupiah(net_income),
            delta=accounting_system.format_rupiah(net_income) if net_income != 0 else None,
            delta_color="normal" if net_income >= 0 else "inverse"
        )
    
    with tab2:
        st.subheader("Laporan Posisi Keuangan (Neraca)")
        st.markdown("---")
        
        st.write("**ASET**")
        if not financials['balance_sheet']['assets'].empty:
            st.dataframe(financials['balance_sheet']['assets'], use_container_width=True, hide_index=True)
            st.write(f"**Total Aset:** {accounting_system.format_rupiah(financials['balance_sheet']['total_assets'])}")
        else:
            st.info("Tidak ada aset")
        
        st.write("**KEWAJIBAN**")
        if not financials['balance_sheet']['liabilities'].empty:
            st.dataframe(financials['balance_sheet']['liabilities'], use_container_width=True, hide_index=True)
            st.write(f"**Total Kewajiban:** {accounting_system.format_rupiah(financials['balance_sheet']['total_liabilities'])}")
        else:
            st.info("Tidak ada kewajiban")
        
        st.write("**MODAL**")
        if not financials['balance_sheet']['equity'].empty:
            st.dataframe(financials['balance_sheet']['equity'], use_container_width=True, hide_index=True)
            st.write(f"**Total Modal:** {accounting_system.format_rupiah(financials['balance_sheet']['total_equity'])}")
        else:
            st.info("Tidak ada modal")
        
        st.markdown("---")
        balance_status = "✅ Balance" if financials['balance_sheet']['total_assets'] == financials['balance_sheet']['total_liabilities_equity'] else "❌ Tidak Balance"
        st.write(f"**Status Neraca:** {balance_status}")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Aset", accounting_system.format_rupiah(financials['balance_sheet']['total_assets']))
        with col2:
            st.metric("Total Kewajiban + Modal", accounting_system.format_rupiah(financials['balance_sheet']['total_liabilities_equity']))
        with col3:
            st.metric("Status", "Balance" if financials['balance_sheet']['total_assets'] == financials['balance_sheet']['total_liabilities_equity'] else "Tidak Balance")
    
    with tab3:
        st.subheader("Download Laporan Keuangan")
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            income_data = {
                'Keterangan': ['Total Pendapatan', 'Total Beban', 'Laba/Rugi Bersih'],
                'Jumlah (Rp)': [
                    accounting_system.format_rupiah(financials['income_statement']['revenue']),
                    accounting_system.format_rupiah(financials['income_statement']['expense']),
                    accounting_system.format_rupiah(financials['income_statement']['net_income'])
                ]
            }
            income_df = pd.DataFrame(income_data)
            excel_income = to_excel(income_df)
            st.download_button(
                label="📥 Download Excel Laba Rugi",
                data=excel_income,
                file_name="laporan_laba_rugi.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )
        
        with col2:
            balance_data = pd.concat([
                financials['balance_sheet']['assets'],
                financials['balance_sheet']['liabilities'], 
                financials['balance_sheet']['equity']
            ], ignore_index=True)
            excel_balance = to_excel(balance_data)
            st.download_button(
                label="📥 Download Excel Neraca",
                data=excel_balance,
                file_name="laporan_neraca.xlsx",
                mime="application/vnd.ms-excel",
                use_container_width=True
            )

# ===== FUNGSI LOGIN SECTION =====
def login_section():
    """Tampilan login dan registrasi"""
    st.title("🔐 Sistem Akuntansi - Desa Wisata Kandri Semarang")
    st.markdown("---")
    
    # Initialize UserManager
    if 'user_manager' not in st.session_state:
        st.session_state.user_manager = UserManager()
    
    user_manager = st.session_state.user_manager
    
    # Tab untuk Login dan Daftar
    tab1, tab2 = st.tabs(["🚀 Login", "📝 Daftar Akun Baru"])
    
    with tab1:
        st.header("Login ke Sistem")
        
        with st.form("login_form"):
            email = st.text_input("📧 Email", placeholder="contoh: nama@gmail.com atau nama@yahoo.com")
            password = st.text_input("🔒 Password", type="password", placeholder="Masukkan password Anda")
            
            login_button = st.form_submit_button("🚀 Login")
            
            if login_button:
                if not email or not password:
                    st.error("❌ Email dan password harus diisi!")
                elif not validate_email(email):
                    st.error("❌ Hanya email @gmail.com atau @yahoo.com yang diperbolehkan!")
                else:
                    success, message = user_manager.verify_user(email, password)
                    if success:
                        st.session_state.logged_in = True
                        st.session_state.user_email = email
                        st.success(f"✅ {message} Selamat datang {email}!")
                        st.balloons()
                        st.rerun()
                    else:
                        st.error(f"❌ {message}")
    
    with tab2:
        st.header("Buat Akun Baru")
        
        with st.form("register_form"):
            new_email = st.text_input("📧 Email Baru", placeholder="contoh: nama@gmail.com atau nama@yahoo.com")
            new_password = st.text_input("🔒 Password Baru", type="password", placeholder="Minimal 6 karakter")
            confirm_password = st.text_input("🔒 Konfirmasi Password", type="password", placeholder="Ketik ulang password")
            
            register_button = st.form_submit_button("📝 Daftar Akun")
            
            if register_button:
                if not new_email or not new_password or not confirm_password:
                    st.error("❌ Semua field harus diisi!")
                elif not validate_email(new_email):
                    st.error("❌ Hanya email @gmail.com atau @yahoo.com yang diperbolehkan!")
                else:
                    # Validasi password
                    is_valid, msg = validate_password(new_password)
                    if not is_valid:
                        st.error(f"❌ {msg}")
                    elif new_password != confirm_password:
                        st.error("❌ Password dan konfirmasi password tidak sama!")
                    else:
                        success, message = user_manager.register_user(new_email, new_password)
                        if success:
                            st.success(f"✅ {message} Silakan login dengan akun Anda.")
                        else:
                            st.error(f"❌ {message}")
    
    # Informasi
    st.markdown("---")
    st.info("**💡 Tips:** Gunakan email @gmail.com atau @yahoo.com dengan password minimal 6 karakter")

def main_app():
    """Aplikasi utama setelah login"""
    
    if 'accounting_system' not in st.session_state:
        st.session_state.accounting_system = AccountingSystem()
    
    accounting_system = st.session_state.accounting_system
    
    st.sidebar.title("🏞️ Desa Wisata Kandri")
    st.sidebar.write(f"👤 **User:** {st.session_state.user_email}")
    st.sidebar.markdown("---")
    
    st.sidebar.header("📊 SIKLUS AKUNTANSI")
    
    menu_options = [
        "📋 Chart of Accounts (COA)",
        "📥 Input Transaksi", 
        "📔 Jurnal Umum",
        "📒 Buku Besar", 
        "⚖️ Neraca Saldo",
        "🔄 Jurnal Penyesuaian",
        "⚖️ Neraca Saldo Penyesuaian",
        "📈 Laporan Keuangan"
    ]
    
    selected_menu = st.sidebar.radio("Pilih Menu:", menu_options)
    
    st.sidebar.markdown("---")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.user_email = ""
        st.rerun()
    
    st.sidebar.markdown("---")
    st.sidebar.header("💡 Info Cepat")
    
    total_transactions = len(accounting_system.transactions)
    total_adjustments = len(accounting_system.adjustments)
    
    st.sidebar.metric("Total Transaksi", total_transactions)
    st.sidebar.metric("Jurnal Penyesuaian", total_adjustments)
    
    st.title(f"{selected_menu.split(' ')[1]} - Sistem Akuntansi")
    st.markdown("---")
    
    if "COA" in selected_menu:
        show_coa_page(accounting_system)
    elif "Input Transaksi" in selected_menu:
        show_input_transaksi_page(accounting_system)
    elif "Jurnal Umum" in selected_menu:
        show_jurnal_umum_page(accounting_system)
    elif "Buku Besar" in selected_menu:
        show_buku_besar_page(accounting_system)
    elif "Neraca Saldo" in selected_menu and "Penyesuaian" not in selected_menu:
        show_neraca_saldo_page(accounting_system)
    elif "Penyesuaian" in selected_menu and "Neraca" not in selected_menu:
        show_penyesuaian_page(accounting_system)
    elif "Neraca Saldo Penyesuaian" in selected_menu:
        show_neraca_saldo_penyesuaian_page(accounting_system)
    elif "Laporan Keuangan" in selected_menu:
        show_laporan_keuangan_page(accounting_system)

def main():
    st.set_page_config(
        page_title="Sistem Akuntansi - Desa Wisata Kandri Semarang",
        page_icon="🏞️",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    if 'logged_in' not in st.session_state:
        st.session_state.logged_in = False
    if 'user_email' not in st.session_state:
        st.session_state.user_email = ""
    
    if not st.session_state.logged_in:
        login_section()
    else:
        main_app()

if __name__ == "__main__":
    main()