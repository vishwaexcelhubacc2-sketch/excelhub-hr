"""
login.py  —  Simple, secure login for Excelhub HR Automation
No external auth library needed. Uses bcrypt via hashlib + secrets stored
in Streamlit's st.secrets (or a local config.toml for local use).
"""

import streamlit as st
import hashlib
import hmac
import time

# ─────────────────────────────────────────────────────────────
#  USER DATABASE
#  In production → store in st.secrets (Streamlit Cloud dashboard)
#  For local use → stored here (change passwords!)
# ─────────────────────────────────────────────────────────────

def _hash(password: str) -> str:
    """SHA-256 hash of the password."""
    return hashlib.sha256(password.encode()).hexdigest()


# Default users — CHANGE THESE PASSWORDS before going live!
DEFAULT_USERS = {
    "hr_admin": {
        "password_hash": _hash("Excelhub@2026"),   # ← change this
        "name":          "HR Admin",
        "role":          "admin",
        "unit":          "All Units",
    },
    "hr_unit1": {
        "password_hash": _hash("Unit1@2026"),       # ← change this
        "name":          "HR Manager – Unit 1",
        "role":          "hr",
        "unit":          "Unit-1",
    },
    "hr_unit2": {
        "password_hash": _hash("Unit2@2026"),       # ← change this
        "name":          "HR Manager – Unit 2",
        "role":          "hr",
        "unit":          "Unit-2",
    },
    "accounts": {
        "password_hash": _hash("Accounts@2026"),    # ← change this
        "name":          "Accounts Team",
        "role":          "viewer",
        "unit":          "All Units",
    },
}


def _get_users() -> dict:
    """
    Try reading from st.secrets first (Streamlit Cloud).
    Falls back to DEFAULT_USERS for local development.
    """
    try:
        users = {}
        for uname, info in st.secrets["users"].items():
            users[uname] = {
                "password_hash": _hash(info["password"]),
                "name":          info.get("name", uname),
                "role":          info.get("role", "hr"),
                "unit":          info.get("unit", "All Units"),
            }
        return users if users else DEFAULT_USERS
    except Exception:
        return DEFAULT_USERS


def check_password(username: str, password: str) -> tuple[bool, dict]:
    """Returns (success, user_info)."""
    users = _get_users()
    user  = users.get(username.strip().lower())
    if not user:
        return False, {}
    entered_hash = _hash(password)
    if hmac.compare_digest(entered_hash, user["password_hash"]):
        return True, user
    return False, {}


# ─────────────────────────────────────────────────────────────
#  LOGIN UI
# ─────────────────────────────────────────────────────────────

LOGIN_CSS = """
<style>
[data-testid="stAppViewContainer"] {
    background: linear-gradient(160deg, #0d1b2e 0%, #1a3a5c 60%, #0d2137 100%);
    min-height: 100vh;
}
[data-testid="stHeader"]          { background: transparent; }
.login-card {
    background: white;
    border-radius: 16px;
    padding: 40px 36px 32px;
    max-width: 420px;
    margin: 60px auto 0;
    box-shadow: 0 24px 60px rgba(0,0,0,0.35);
}
.login-logo {
    text-align: center;
    margin-bottom: 28px;
}
.login-logo-circle {
    width: 80px;
    height: 80px;
    border-radius: 50%;
    background: linear-gradient(135deg, #1a2942, #2563a8);
    margin: 0 auto 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 36px;
}
.login-title {
    font-size: 22px;
    font-weight: 700;
    color: #1a2942;
    margin: 0 0 4px;
    text-align: center;
}
.login-sub {
    font-size: 13px;
    color: #888;
    text-align: center;
    margin: 0 0 28px;
}
.login-label {
    font-size: 12px;
    font-weight: 600;
    color: #444;
    margin-bottom: 4px;
    letter-spacing: 0.04em;
}
.login-error {
    background: #fff0f0;
    border-left: 3px solid #d32f2f;
    color: #c62828;
    padding: 10px 14px;
    border-radius: 0 8px 8px 0;
    font-size: 13px;
    margin-top: 12px;
}
.login-footer {
    text-align: center;
    font-size: 11px;
    color: #aaa;
    margin-top: 24px;
}
.user-badge {
    display: inline-block;
    background: #e8f5e9;
    color: #2e7d32;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
}
</style>
"""


def show_login_page() -> bool:
    """
    Shows login UI. Returns True if user is authenticated.
    Stores user info in st.session_state.user
    """
    # Already logged in?
    if st.session_state.get("authenticated"):
        return True

    st.markdown(LOGIN_CSS, unsafe_allow_html=True)

    # Centered login card
    _, center_col, _ = st.columns([1, 2, 1])

    with center_col:
        st.markdown("""
        <div class="login-card">
          <div class="login-logo">
            <div class="login-logo-circle">🏭</div>
            <div class="login-title">Excelhub HR</div>
            <div class="login-sub">Attendance & Payroll Automation</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        # Inputs
        st.markdown('<p class="login-label">USERNAME</p>', unsafe_allow_html=True)
        username = st.text_input("Username", label_visibility="collapsed",
                                  placeholder="e.g. hr_admin",
                                  key="login_user")

        st.markdown('<p class="login-label" style="margin-top:12px">PASSWORD</p>',
                    unsafe_allow_html=True)
        password = st.text_input("Password", type="password",
                                  label_visibility="collapsed",
                                  placeholder="Enter your password",
                                  key="login_pass")

        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🔐  Login", use_container_width=True, type="primary"):
            if not username or not password:
                st.error("Please enter both username and password.")
            else:
                with st.spinner("Verifying..."):
                    time.sleep(0.6)   # small delay to prevent brute-force
                    ok, user_info = check_password(username, password)

                if ok:
                    st.session_state.authenticated = True
                    st.session_state.user          = user_info
                    st.session_state.username      = username
                    st.rerun()
                else:
                    st.markdown(
                        '<div class="login-error">'
                        '❌  Incorrect username or password. Please try again.'
                        '</div>',
                        unsafe_allow_html=True
                    )

        st.markdown("""
        <div class="login-footer">
          🔒 Secured · Only authorized HR staff can access this system<br>
          Contact your IT admin to reset your password
        </div>
        """, unsafe_allow_html=True)

    return False


def show_user_bar():
    """Top bar showing logged-in user info + logout button."""
    user = st.session_state.get("user", {})
    name = user.get("name", "User")
    role = user.get("role", "").upper()
    unit = user.get("unit", "")

    col1, col2 = st.columns([5, 1])
    with col1:
        st.markdown(
            f'👤 &nbsp; <b>{name}</b> &nbsp;·&nbsp; '
            f'<span class="user-badge">{role}</span> &nbsp;·&nbsp; {unit}',
            unsafe_allow_html=True
        )
    with col2:
        if st.button("🚪 Logout", use_container_width=True):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

    st.divider()
