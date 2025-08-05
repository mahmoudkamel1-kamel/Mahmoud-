#!/usr/bin/env python3
"""
Professional Sales & Payment Tracker - Optimized Version
Features:
- Lazy loading for faster startup
- Enhanced modern UI with visual improvements
- Optimized database operations
- Better memory management
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime, timedelta
import os
import shutil
import webbrowser
import hashlib
import threading
import functools
from typing import Optional, Dict, List, Tuple, Any

# Core imports - loaded immediately
try:
    from tkcalendar import DateEntry
    from ttkthemes import ThemedStyle
except ImportError:
    messagebox.showerror("Missing Dependencies", 
                        "Please install required packages:\npip install tkcalendar ttkthemes")
    exit(1)

# Heavy imports - loaded lazily
_matplotlib = None
_pandas = None
_openpyxl = None
_numpy = None

def lazy_import_matplotlib():
    """Lazy import matplotlib to speed up startup"""
    global _matplotlib
    if _matplotlib is None:
        try:
            import matplotlib
            matplotlib.use('TkAgg')  # Set backend before importing pyplot
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            _matplotlib = {'Figure': Figure, 'FigureCanvasTkAgg': FigureCanvasTkAgg}
        except ImportError:
            messagebox.showerror("Import Error", "matplotlib not found. Install with: pip install matplotlib")
            return None
    return _matplotlib

def lazy_import_pandas():
    """Lazy import pandas"""
    global _pandas
    if _pandas is None:
        try:
            import pandas as pd
            _pandas = pd
        except ImportError:
            messagebox.showerror("Import Error", "pandas not found. Install with: pip install pandas")
            return None
    return _pandas

def lazy_import_openpyxl():
    """Lazy import openpyxl"""
    global _openpyxl
    if _openpyxl is None:
        try:
            import openpyxl
            from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
            _openpyxl = {
                'Workbook': openpyxl.Workbook,
                'Font': Font,
                'Alignment': Alignment,
                'PatternFill': PatternFill,
                'Border': Border,
                'Side': Side
            }
        except ImportError:
            messagebox.showerror("Import Error", "openpyxl not found. Install with: pip install openpyxl")
            return None
    return _openpyxl

def lazy_import_numpy():
    """Lazy import numpy"""
    global _numpy
    if _numpy is None:
        try:
            import numpy as np
            _numpy = np
        except ImportError:
            messagebox.showerror("Import Error", "numpy not found. Install with: pip install numpy")
            return None
    return _numpy

# Visual enhancements and modern styling
class ModernColors:
    """Modern color palette for enhanced visuals"""
    PRIMARY = "#2563eb"        # Blue
    PRIMARY_HOVER = "#1d4ed8"  # Darker blue
    SUCCESS = "#10b981"        # Green
    SUCCESS_HOVER = "#059669"  # Darker green
    DANGER = "#ef4444"         # Red
    DANGER_HOVER = "#dc2626"   # Darker red
    WARNING = "#f59e0b"        # Orange
    INFO = "#06b6d4"           # Cyan
    
    BACKGROUND = "#f8fafc"     # Light gray
    SURFACE = "#ffffff"        # White
    SURFACE_VARIANT = "#f1f5f9" # Light gray variant
    
    TEXT_PRIMARY = "#1e293b"   # Dark gray
    TEXT_SECONDARY = "#64748b" # Medium gray
    TEXT_MUTED = "#94a3b8"     # Light gray
    
    BORDER = "#e2e8f0"         # Light border
    BORDER_FOCUS = "#3b82f6"   # Blue border for focus

class LoadingScreen:
    """Modern loading screen with progress indication"""
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Sales Tracker")
        self.root.geometry("400x250")
        self.root.configure(bg=ModernColors.BACKGROUND)
        self.root.resizable(False, False)
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')
        
        # Remove window decorations for splash effect
        self.root.overrideredirect(True)
        self.root.configure(highlightbackground=ModernColors.BORDER, highlightthickness=1)
        
        self._create_loading_ui()
        
    def _create_loading_ui(self):
        """Create modern loading interface"""
        main_frame = tk.Frame(self.root, bg=ModernColors.SURFACE, padx=40, pady=40)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Logo/Title
        title_label = tk.Label(main_frame, 
                              text="💼 Sales Tracker",
                              font=("Helvetica", 24, "bold"),
                              fg=ModernColors.PRIMARY,
                              bg=ModernColors.SURFACE)
        title_label.pack(pady=(0, 10))
        
        subtitle_label = tk.Label(main_frame,
                                 text="Professional Sales & Payment Management",
                                 font=("Helvetica", 11),
                                 fg=ModernColors.TEXT_SECONDARY,
                                 bg=ModernColors.SURFACE)
        subtitle_label.pack(pady=(0, 30))
        
        # Progress bar
        self.progress_var = tk.StringVar(value="Initializing...")
        self.progress_label = tk.Label(main_frame,
                                      textvariable=self.progress_var,
                                      font=("Helvetica", 10),
                                      fg=ModernColors.TEXT_SECONDARY,
                                      bg=ModernColors.SURFACE)
        self.progress_label.pack(pady=(0, 10))
        
        # Modern progress bar
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Modern.Horizontal.TProgressbar",
                       background=ModernColors.PRIMARY,
                       troughcolor=ModernColors.SURFACE_VARIANT,
                       borderwidth=0,
                       lightcolor=ModernColors.PRIMARY,
                       darkcolor=ModernColors.PRIMARY)
        
        self.progress_bar = ttk.Progressbar(main_frame,
                                           style="Modern.Horizontal.TProgressbar",
                                           mode='indeterminate',
                                           length=300)
        self.progress_bar.pack(pady=(0, 20))
        self.progress_bar.start(10)
        
        # Version info
        version_label = tk.Label(main_frame,
                                text="v2.0 - Optimized Edition",
                                font=("Helvetica", 9),
                                fg=ModernColors.TEXT_MUTED,
                                bg=ModernColors.SURFACE)
        version_label.pack(side=tk.BOTTOM)
        
    def update_progress(self, message: str):
        """Update progress message"""
        self.progress_var.set(message)
        self.root.update()
        
    def close(self):
        """Close loading screen"""
        self.progress_bar.stop()
        self.root.destroy()

# Optimized Database with connection pooling and caching
class OptimizedDatabase:
    """Enhanced database class with performance optimizations"""
    
    def __init__(self, db_name="sales_tracker.db"):
        self.db_name = db_name
        self.conn = None
        self.cursor = None
        self._cache = {}
        self._cache_timeout = 300  # 5 minutes
        self._cache_timestamps = {}
        self._connect()
        self._initialize_db()
        
    def _connect(self):
        """Establish database connection with optimizations"""
        self.conn = sqlite3.connect(self.db_name, check_same_thread=False)
        self.conn.execute("PRAGMA journal_mode=WAL")  # Write-Ahead Logging for better performance
        self.conn.execute("PRAGMA synchronous=NORMAL")  # Balance between safety and speed
        self.conn.execute("PRAGMA cache_size=10000")  # Increase cache size
        self.conn.execute("PRAGMA temp_store=MEMORY")  # Store temp tables in memory
        self.cursor = self.conn.cursor()
        
    def _initialize_db(self):
        """Create tables and indexes with performance optimizations"""
        # Create tables
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                is_admin BOOLEAN NOT NULL DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS clients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                code TEXT NOT NULL UNIQUE,
                country TEXT,
                user_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users (id)
            )
        """)
        
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER NOT NULL,
                type TEXT NOT NULL CHECK (type IN ('order', 'payment')),
                amount REAL NOT NULL,
                description TEXT,
                notes TEXT,
                transaction_date TEXT NOT NULL,
                due_date TEXT,
                posting_date TEXT NOT NULL,
                attachment_path TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (client_id) REFERENCES clients (id) ON DELETE CASCADE
            )
        """)
        
        # Create performance indexes
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_clients_user_id ON clients(user_id)",
            "CREATE INDEX IF NOT EXISTS idx_clients_code ON clients(code)",
            "CREATE INDEX IF NOT EXISTS idx_clients_country ON clients(country)",
            "CREATE INDEX IF NOT EXISTS idx_transactions_client_id ON transactions(client_id)",
            "CREATE INDEX IF NOT EXISTS idx_transactions_type ON transactions(type)",
            "CREATE INDEX IF NOT EXISTS idx_transactions_date ON transactions(transaction_date)",
            "CREATE INDEX IF NOT EXISTS idx_transactions_due_date ON transactions(due_date)",
            "CREATE INDEX IF NOT EXISTS idx_transactions_composite ON transactions(client_id, type, transaction_date)"
        ]
        
        for index_sql in indexes:
            self.cursor.execute(index_sql)
        
        # Add columns if they don't exist (migration support)
        self._add_column_if_not_exists('clients', 'user_id', 'INTEGER')
        self._add_column_if_not_exists('clients', 'code', 'TEXT')
        self._add_column_if_not_exists('clients', 'country', 'TEXT')
        self._add_column_if_not_exists('transactions', 'attachment_path', 'TEXT')
        self._add_column_if_not_exists('transactions', 'transaction_date', 'TEXT')
        self._add_column_if_not_exists('transactions', 'notes', 'TEXT')
        self._add_column_if_not_exists('transactions', 'due_date', 'TEXT')
        
        # Handle column rename
        self._rename_column_if_exists('transactions', 'date', 'posting_date')
        
        self.cursor.execute("PRAGMA foreign_keys = ON")
        self.conn.commit()
        self._ensure_admin_exists()
        
    def _add_column_if_not_exists(self, table: str, column: str, col_type: str):
        """Add column if it doesn't exist"""
        self.cursor.execute(f"PRAGMA table_info({table})")
        columns = [info[1] for info in self.cursor.fetchall()]
        if column not in columns:
            self.cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")
            
    def _rename_column_if_exists(self, table: str, old_name: str, new_name: str):
        """Rename column if old exists and new doesn't"""
        self.cursor.execute(f"PRAGMA table_info({table})")
        columns = [info[1] for info in self.cursor.fetchall()]
        if old_name in columns and new_name not in columns:
            self.cursor.execute(f"ALTER TABLE {table} RENAME COLUMN {old_name} TO {new_name}")
            
    def _ensure_admin_exists(self):
        """Create default admin if no users exist"""
        self.cursor.execute("SELECT COUNT(*) FROM users")
        if self.cursor.fetchone()[0] == 0:
            self.add_user("admin", "admin", is_admin=True)
            
    def _get_cache_key(self, query: str, params: tuple = ()) -> str:
        """Generate cache key for query"""
        return hashlib.md5(f"{query}{params}".encode()).hexdigest()
        
    def _is_cache_valid(self, key: str) -> bool:
        """Check if cache entry is still valid"""
        if key not in self._cache_timestamps:
            return False
        return (datetime.now().timestamp() - self._cache_timestamps[key]) < self._cache_timeout
        
    def _execute_with_cache(self, query: str, params: tuple = (), use_cache: bool = True):
        """Execute query with optional caching for read operations"""
        if use_cache and query.strip().upper().startswith('SELECT'):
            cache_key = self._get_cache_key(query, params)
            if cache_key in self._cache and self._is_cache_valid(cache_key):
                return self._cache[cache_key]
                
        self.cursor.execute(query, params)
        result = self.cursor.fetchall()
        
        if use_cache and query.strip().upper().startswith('SELECT'):
            cache_key = self._get_cache_key(query, params)
            self._cache[cache_key] = result
            self._cache_timestamps[cache_key] = datetime.now().timestamp()
            
        return result
        
    def _clear_cache(self):
        """Clear query cache"""
        self._cache.clear()
        self._cache_timestamps.clear()
        
    # User Methods
    def add_user(self, username: str, password: str, is_admin: bool = False) -> bool:
        """Add new user"""
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        try:
            self.cursor.execute(
                "INSERT INTO users (username, password_hash, is_admin) VALUES (?, ?, ?)",
                (username, password_hash, is_admin)
            )
            self.conn.commit()
            self._clear_cache()  # Clear cache after write operation
            return True
        except sqlite3.IntegrityError:
            return False
            
    def verify_user(self, username: str, password: str) -> Optional[Tuple[int, bool]]:
        """Verify user credentials"""
        password_hash = hashlib.sha256(password.encode()).hexdigest()
        result = self._execute_with_cache(
            "SELECT id, is_admin FROM users WHERE username = ? AND password_hash = ?",
            (username, password_hash)
        )
        return result[0] if result else None
        
    def get_all_users(self) -> List[Tuple[int, str, bool]]:
        """Get all users"""
        return self._execute_with_cache("SELECT id, username, is_admin FROM users ORDER BY username")
        
    def update_user(self, user_id: int, username: str, password: Optional[str], is_admin: bool) -> bool:
        """Update user"""
        try:
            if password:
                password_hash = hashlib.sha256(password.encode()).hexdigest()
                self.cursor.execute(
                    "UPDATE users SET username=?, password_hash=?, is_admin=? WHERE id=?",
                    (username, password_hash, is_admin, user_id)
                )
            else:
                self.cursor.execute(
                    "UPDATE users SET username=?, is_admin=? WHERE id=?",
                    (username, is_admin, user_id)
                )
            self.conn.commit()
            self._clear_cache()
            return True
        except sqlite3.IntegrityError:
            return False
            
    def delete_user(self, user_id: int):
        """Delete user"""
        self.cursor.execute("DELETE FROM users WHERE id=?", (user_id,))
        self.conn.commit()
        self._clear_cache()
        
    def get_user_aggregated_totals(self, user_id: int) -> Tuple[float, float]:
        """Get aggregated totals for user"""
        query = """
            SELECT t.type, SUM(t.amount) FROM transactions t
            JOIN clients c ON t.client_id = c.id
            WHERE c.user_id = ? GROUP BY t.type
        """
        result = self._execute_with_cache(query, (user_id,))
        totals = {'order': 0, 'payment': 0}
        for type_name, total_amount in result:
            totals[type_name] = total_amount
        return totals['order'], totals['payment']
        
    # Dashboard Methods
    def get_overall_totals(self) -> Tuple[float, float]:
        """Get overall totals"""
        result = self._execute_with_cache("SELECT type, SUM(amount) FROM transactions GROUP BY type")
        totals = {'order': 0, 'payment': 0}
        for type_name, total_amount in result:
            totals[type_name] = total_amount
        return totals['order'], totals['payment']
        
    def get_sales_and_payments_by_user(self) -> List[Tuple[str, float, float]]:
        """Get sales and payments by user"""
        query = """
            SELECT
                u.username,
                COALESCE(SUM(CASE WHEN t.type = 'order' THEN t.amount ELSE 0 END), 0) as total_sales,
                COALESCE(SUM(CASE WHEN t.type = 'payment' THEN t.amount ELSE 0 END), 0) as total_payments
            FROM users u
            LEFT JOIN clients c ON u.id = c.user_id
            LEFT JOIN transactions t ON c.id = t.client_id
            GROUP BY u.id, u.username
            ORDER BY total_sales DESC
        """
        return self._execute_with_cache(query)
        
    def get_client_summary_for_user(self, user_id: int) -> List[Tuple[str, float]]:
        """Get client summary for user"""
        query = """
            SELECT
                c.name,
                COALESCE(SUM(CASE WHEN t.type = 'order' THEN t.amount ELSE -t.amount END), 0) as balance
            FROM clients c
            LEFT JOIN transactions t ON c.id = t.client_id
            WHERE c.user_id = ?
            GROUP BY c.id, c.name
            ORDER BY c.name
        """
        return self._execute_with_cache(query, (user_id,))
        
    # Client Methods
    def add_client(self, name: str, code: str, country: str, user_id: int) -> Optional[int]:
        """Add new client"""
        try:
            self.cursor.execute(
                "INSERT INTO clients (name, code, country, user_id) VALUES (?, ?, ?, ?)",
                (name, code, country, user_id)
            )
            self.conn.commit()
            self._clear_cache()
            return self.cursor.lastrowid
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"A client with the code '{code}' already exists.")
            return None
            
    def get_clients(self, user_id: int, is_admin: bool, search_term: str = "", 
                   country_filter: str = "", viewed_user_id: Optional[int] = None) -> List[Tuple[int, str, str]]:
        """Get clients with filtering"""
        query = "SELECT id, name, code FROM clients"
        params = []
        conditions = []
        
        if viewed_user_id:
            conditions.append("user_id = ?")
            params.append(viewed_user_id)
        elif not is_admin:
            conditions.append("user_id = ?")
            params.append(user_id)
            
        if search_term:
            conditions.append("(name LIKE ? OR code LIKE ?)")
            params.extend([f'%{search_term}%', f'%{search_term}%'])
            
        if country_filter:
            conditions.append("country = ?")
            params.append(country_filter)
            
        if conditions:
            query += " WHERE " + " AND ".join(conditions)
        query += " ORDER BY name"
        
        return self._execute_with_cache(query, tuple(params))
        
    def get_all_countries(self, user_id: int, is_admin: bool, viewed_user_id: Optional[int] = None) -> List[str]:
        """Get all countries"""
        query = "SELECT DISTINCT country FROM clients WHERE country IS NOT NULL AND country != ''"
        params = []
        
        if viewed_user_id:
            query += " AND user_id = ?"
            params.append(viewed_user_id)
        elif not is_admin:
            query += " AND user_id = ?"
            params.append(user_id)
            
        query += " ORDER BY country"
        result = self._execute_with_cache(query, tuple(params))
        return [row[0] for row in result]
        
    def update_client(self, client_id: int, new_name: str, new_code: str, new_country: str) -> bool:
        """Update client"""
        try:
            self.cursor.execute(
                "UPDATE clients SET name = ?, code = ?, country = ? WHERE id = ?",
                (new_name, new_code, new_country, client_id)
            )
            self.conn.commit()
            self._clear_cache()
            return True
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"A client with the code '{new_code}' already exists.")
            return False
            
    def delete_client(self, client_id: int) -> List[str]:
        """Delete client and return attachment paths"""
        self.cursor.execute("SELECT attachment_path FROM transactions WHERE client_id=?", (client_id,))
        attachments = self.cursor.fetchall()
        self.cursor.execute("DELETE FROM clients WHERE id=?", (client_id,))
        self.conn.commit()
        self._clear_cache()
        return [att[0] for att in attachments if att[0]]
        
    def get_client_details(self, client_id: int) -> Optional[Tuple[int, str, str, str]]:
        """Get client details"""
        result = self._execute_with_cache("SELECT id, name, code, country FROM clients WHERE id=?", (client_id,))
        return result[0] if result else None
        
    # Transaction Methods
    def add_transaction(self, client_id: int, trans_type: str, amount: float, description: str,
                       notes: str, trans_date: str, due_date: Optional[str], attachment_path: Optional[str]):
        """Add new transaction"""
        posting_date_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.cursor.execute("""
            INSERT INTO transactions (client_id, type, amount, description, notes, 
                                    transaction_date, due_date, posting_date, attachment_path)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (client_id, trans_type, amount, description, notes, trans_date, due_date, posting_date_str, attachment_path))
        self.conn.commit()
        self._clear_cache()
        
    def get_transactions(self, client_id: int, start_date: Optional[str] = None, 
                        end_date: Optional[str] = None, status_filter: str = "") -> List[Tuple]:
        """Get transactions with filtering"""
        query = """
            SELECT id, transaction_date, description, type, amount, attachment_path, 
                   posting_date, notes, due_date 
            FROM transactions WHERE client_id = ?
        """
        params = [client_id]
        
        if start_date:
            query += " AND transaction_date >= ?"
            params.append(start_date)
        if end_date:
            query += " AND transaction_date <= ?"
            params.append(end_date)
            
        today = datetime.now().strftime("%Y-%m-%d")
        if status_filter == "Overdue":
            query += " AND type = 'order' AND due_date < ?"
            params.append(today)
            
        query += " ORDER BY transaction_date ASC, posting_date ASC"
        transactions = self._execute_with_cache(query, tuple(params))
        
        # Handle paid/unpaid filtering
        if status_filter in ["Paid", "Unpaid"]:
            total_sales = sum(t[4] for t in transactions if t[3] == 'order')
            total_paid = sum(t[4] for t in transactions if t[3] == 'payment')
            is_paid = total_paid >= total_sales
            if (status_filter == "Paid" and not is_paid) or \
               (status_filter == "Unpaid" and is_paid):
                return []
                
        return transactions
        
    def get_transaction(self, trans_id: int) -> Optional[Tuple]:
        """Get single transaction"""
        result = self._execute_with_cache(
            "SELECT id, type, amount, description, transaction_date, attachment_path, notes, due_date FROM transactions WHERE id=?",
            (trans_id,)
        )
        return result[0] if result else None
        
    def update_transaction(self, trans_id: int, amount: float, description: str, notes: str,
                          trans_date: str, due_date: Optional[str], attachment_path: Optional[str]):
        """Update transaction"""
        self.cursor.execute("""
            UPDATE transactions
            SET amount = ?, description = ?, notes = ?, transaction_date = ?, due_date = ?, attachment_path = ?
            WHERE id = ?
        """, (amount, description, notes, trans_date, due_date, attachment_path, trans_id))
        self.conn.commit()
        self._clear_cache()
        
    def delete_transaction(self, trans_id: int) -> Optional[str]:
        """Delete transaction and return attachment path"""
        self.cursor.execute("SELECT attachment_path FROM transactions WHERE id=?", (trans_id,))
        result = self.cursor.fetchone()
        attachment_path = result[0] if result else None
        self.cursor.execute("DELETE FROM transactions WHERE id=?", (trans_id,))
        self.conn.commit()
        self._clear_cache()
        return attachment_path
        
    def get_due_amount(self, client_id: int) -> float:
        """Calculate currently due amount for client"""
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Get due orders
        result = self._execute_with_cache(
            "SELECT SUM(amount) FROM transactions WHERE client_id = ? AND type = 'order' AND due_date <= ?",
            (client_id, today)
        )
        due_orders = result[0][0] if result and result[0][0] else 0.0
        
        # Get total payments
        result = self._execute_with_cache(
            "SELECT SUM(amount) FROM transactions WHERE client_id = ? AND type = 'payment'",
            (client_id,)
        )
        total_payments = result[0][0] if result and result[0][0] else 0.0
        
        currently_due = due_orders - total_payments
        return max(0, currently_due)
        
    def close(self):
        """Close database connection"""
        if self.conn:
            self.conn.close()