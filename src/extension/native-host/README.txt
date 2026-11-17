PHX Classic Outlook helper

1. Run install-host.cmd. It copies this folder to %LOCALAPPDATA%\PHX\OutlookHost\<version> and registers com.phx.outlook for Chrome and Edge under your user account.
2. Return to PHX Slide Sheet and click Recheck.
3. repair-host.cmd rewrites the current version. uninstall-host.cmd removes the registration and the helper folder.
4. The previous version folder is kept so Install can point the registration back at it with -Action Rollback.

The helper opens a Classic Outlook compose. It does not send mail. New Outlook is not used.
