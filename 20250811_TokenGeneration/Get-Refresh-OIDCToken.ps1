
<#
.SYNOPSIS
    Acquires an OIDC access token using pwps_dab Get-OIDCToken and refreshes it automatically before it expires.

.DESCRIPTION
    This script wraps the pwps_dab module's Get-OIDCToken command (or a similarly named function) and
    maintains a current, auto‑refreshed token in memory. It schedules a refresh a configurable number
    of seconds (skew) before the reported expiration time.

    It exposes:
      Start-OIDCRefreshLoop  - Begins token acquisition & automatic refresh
      Stop-OIDCRefreshLoop   - Stops the refresh timer
      Get-CurrentOIDCToken   - Returns the current token object
      Get-AuthorizationHeader - Convenience helper to build a Bearer header
      Invoke-OIDCRefresh     - Forces an immediate refresh (with optional bypass of backoff)

    Refresh logic:
      - Uses a System.Timers.Timer set to (ExpiresOn - Now - RefreshSkew).
      - If calculated interval < 30s, a minimum 30s interval is used to avoid thrash.
      - Includes simple exponential backoff on transient failures (up to a max backoff window).
      - Thread‑safe with a SyncRoot lock.

    NOTE: Adjust parameter names passed to Get-OIDCToken if your pwps_dab version uses different ones.
          (See the placeholder section marked "ADJUST THIS CALL AS NEEDED".)

.PARAMETER Authority
    The authority / tenant / issuer base URL (e.g. https://login.microsoftonline.com/<tenantId>).

.PARAMETER ClientId
    Client/Application ID used for token acquisition.

.PARAMETER ClientSecret
    (Optional) Client secret, if using a confidential client flow.

.PARAMETER Scope
    Space or comma separated scopes (some modules may expect resource/.default).

.PARAMETER RefreshSkewSeconds
    Number of seconds BEFORE reported expiration to attempt refresh (default 300).

.PARAMETER MinBackoffSeconds
    Starting backoff after a transient failure (default 5).

.PARAMETER MaxBackoffSeconds
    Maximum backoff window (default 300).

.PARAMETER VerboseLogging
    Switch to enable
