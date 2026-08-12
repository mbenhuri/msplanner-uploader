"""
Delegated (authorization code flow) auth for the web app.

Follows the shape of Microsoft's ms-identity-python-webapp sample:
initiate_auth_code_flow() to start the redirect, acquire_token_by_auth_code_flow()
on the callback, acquire_token_silent() for everything afterwards.

Device code flow is deliberately not used — the tenant has it disabled, and it
is a known phishing vector. This is a confidential client (Web platform type),
so it authenticates with a client secret and the redirect URI is validated by
Entra.

TOKEN CACHE — READ THIS BEFORE DEPLOYING
----------------------------------------
The MSAL token cache is serialized to a single file (config TOKEN_CACHE_PATH,
mode 0600). That is fine for a single-admin tool on your own machine, but it is
now the most sensitive secret on disk: it holds a refresh token that acts as
you, and redeeming it does not re-run interactive MFA (Conditional Access is
re-evaluated at refresh, but that is a policy check, not a fresh sign-in).

It must never be committed — .gitignore covers it. Before this serves more than
one user it needs to become a per-user, server-side store: Flask-Session backed
by Azure Cache for Redis, cache keyed by the user's object id, with the client
secret moved to Key Vault or replaced by a certificate / managed identity.
"""

import os
import threading

import msal


class AuthError(RuntimeError):
    """Sign-in is required or could not be renewed silently."""


class AuthManager:
    def __init__(self, config):
        self.config = config
        self.cache_path = config["TOKEN_CACHE_PATH"]
        self.scopes = config["SCOPES"]
        self.redirect_uri = config["REDIRECT_URI"]
        # One shared cache instance: request threads and the background worker
        # threads that run imports both touch it, so access is serialized.
        self._lock = threading.RLock()
        self._cache = msal.SerializableTokenCache()
        self._load()

    # -- cache persistence --------------------------------------------------

    def _load(self):
        if os.path.exists(self.cache_path):
            with open(self.cache_path) as f:
                self._cache.deserialize(f.read())

    def _persist(self):
        if not self._cache.has_state_changed:
            return
        tmp = f"{self.cache_path}.tmp"
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as f:
            f.write(self._cache.serialize())
        os.replace(tmp, self.cache_path)

    def forget(self, home_account_id):
        with self._lock:
            app = self._app()
            for account in app.get_accounts():
                if account.get("home_account_id") == home_account_id:
                    app.remove_account(account)
            self._persist()

    # -- msal ---------------------------------------------------------------

    def _app(self):
        return msal.ConfidentialClientApplication(
            self.config["CLIENT_ID"],
            authority=self.config["AUTHORITY"],
            client_credential=self.config["CLIENT_SECRET"],
            token_cache=self._cache,
        )

    def build_auth_code_flow(self):
        with self._lock:
            return self._app().initiate_auth_code_flow(
                self.scopes, redirect_uri=self.redirect_uri
            )

    def complete_auth_code_flow(self, flow, request_args):
        """Returns the MSAL result dict; caller checks for 'error'."""
        with self._lock:
            result = self._app().acquire_token_by_auth_code_flow(flow, request_args)
            self._persist()
        return result

    def _find_account(self, app, home_account_id):
        for account in app.get_accounts():
            if account.get("home_account_id") == home_account_id:
                return account
        return None

    @staticmethod
    def home_account_id_from(result):
        claims = result.get("id_token_claims") or {}
        oid, tid = claims.get("oid"), claims.get("tid")
        if oid and tid:
            return f"{oid}.{tid}"
        return None

    def token_provider(self, home_account_id):
        """
        Returns callable(force_refresh=False) -> access token, for
        planner_core.GraphClient.

        Deliberately closes over the AuthManager and the account id, not over
        the Flask session — imports run on a background thread with no request
        context, and still need to be able to renew a token mid-run.
        """

        def provider(force_refresh=False):
            with self._lock:
                app = self._app()
                account = self._find_account(app, home_account_id)
                if account is None:
                    raise AuthError("No cached account; sign in again.")
                result = app.acquire_token_silent_with_error(
                    self.scopes, account=account, force_refresh=force_refresh
                )
                self._persist()
            if not result or "access_token" not in result:
                detail = (result or {}).get("error_description") or "no token returned"
                raise AuthError(f"Could not renew access silently: {detail}")
            return result["access_token"]

        return provider

    def has_valid_session(self, home_account_id):
        if not home_account_id:
            return False
        with self._lock:
            return self._find_account(self._app(), home_account_id) is not None
