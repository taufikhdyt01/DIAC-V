# modules/auth.py - Modul otentikasi untuk DIAC-V

import pandas as pd
import bcrypt
import os
from config import USERS_DB, ACCESS_LEVELS

class AuthManager:
    def __init__(self):
        self.current_user = None
        self._create_default_db_if_not_exists()
        
    def _create_default_db_if_not_exists(self):
        """Membuat database pengguna default jika belum ada"""
        if not os.path.exists(USERS_DB):
            # Buat direktori jika belum ada
            os.makedirs(os.path.dirname(USERS_DB), exist_ok=True)
            
            # Buat data pengguna default
            default_users = pd.DataFrame({
                'username': ['admin', 'john_ade', 'mary_bdu', 'alex_mar', 'dave_man', 
                             'sarah_prj', 'mike_fin', 'lisa_leg', 'ceo'],
                'password': [self._hash_password('admin123')] + [self._hash_password('password123')] * 8,
                'name': ['Administrator', 'John Smith', 'Mary Johnson', 'Alex Brown', 'Dave Wilson',
                        'Sarah Miller', 'Mike Taylor', 'Lisa Anderson', 'CEO'],
                'department': ['IT', 'ADE', 'BDU', 'MAR', 'MAN', 'PRJ', 'FIN', 'LEG', 'EXEC'],
                'access_level': ['admin', 'user', 'user', 'user', 'user', 'user', 'user', 'user', 'ceo'],
                'email': ['admin@diac-v.com', 'john@diac-v.com', 'mary@diac-v.com', 'alex@diac-v.com', 
                         'dave@diac-v.com', 'sarah@diac-v.com', 'mike@diac-v.com', 'lisa@diac-v.com', 'ceo@diac-v.com']
            })
            
            # Simpan ke Excel
            default_users.to_excel(USERS_DB, index=False)
    
    def _hash_password(self, password):
        """Hash password menggunakan bcrypt"""
        return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    def _verify_password(self, password, hashed):
        """Verifikasi password dengan hash"""
        return bcrypt.checkpw(password.encode('utf-8'), hashed.encode('utf-8'))
    
    def login(self, username, password):
        """Otentikasi pengguna dengan username dan password"""
        # Baca database pengguna
        try:
            df = pd.read_excel(USERS_DB)
            
            # Cari user berdasarkan username
            user = df[df['username'] == username]
            if user.empty:
                return False, "Username tidak ditemukan."
            
            # Verifikasi password
            stored_password = user['password'].values[0]
            if not self._verify_password(password, stored_password):
                return False, "Password salah."
            
            # Set pengguna saat ini
            self.current_user = {
                'username': user['username'].values[0],
                'name': user['name'].values[0],
                'department': user['department'].values[0],
                'access_level': user['access_level'].values[0],
                'email': user['email'].values[0]
            }
            
            return True, "Login berhasil."
        except Exception as e:
            return False, f"Error saat login: {str(e)}"
    
    def logout(self):
        """Logout pengguna saat ini"""
        self.current_user = None
        return True
    
    def get_current_user(self):
        """Dapatkan info pengguna saat ini"""
        return self.current_user
    
    def has_access(self, department):
        """Cek apakah pengguna punya akses ke departemen tertentu"""
        if not self.current_user:
            return False
        
        user_level = ACCESS_LEVELS.get(self.current_user['access_level'], 0)
        user_dept = self.current_user['department']
        
        # Admin dan CEO punya akses ke semua
        if user_level >= ACCESS_LEVELS['admin']:
            return True
        
        # User biasa hanya punya akses ke departemen sendiri
        if user_level == ACCESS_LEVELS['user']:
            return user_dept == department
        
        # Manager dan director punya akses berdasarkan level
        # Logika akses untuk level lain dapat ditambahkan di sini
        
        return False
    
    def get_accessible_departments(self):
        """Dapatkan daftar departemen yang dapat diakses pengguna saat ini"""
        from config import DEPARTMENTS
        
        if not self.current_user:
            return []
        
        # Admin dan CEO dapat mengakses semua departemen
        if self.current_user['access_level'] in ['admin', 'ceo']:
            return [dept["id"] for dept in DEPARTMENTS]
        
        # User biasa hanya dapat mengakses departemen mereka sendiri
        return [self.current_user['department']]
    
    def change_password(self, old_password, new_password):
        """Ubah password pengguna saat ini"""
        if not self.current_user:
            return False, "Tidak ada pengguna yang login."
        
        try:
            df = pd.read_excel(USERS_DB)
            user_idx = df[df['username'] == self.current_user['username']].index[0]
            
            # Verifikasi password lama
            stored_password = df.loc[user_idx, 'password']
            if not self._verify_password(old_password, stored_password):
                return False, "Password lama salah."
            
            # Update password
            df.loc[user_idx, 'password'] = self._hash_password(new_password)
            df.to_excel(USERS_DB, index=False)
            
            return True, "Password berhasil diubah."
        except Exception as e:
            return False, f"Error saat mengubah password: {str(e)}"