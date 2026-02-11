This package allows switching between Always On and General VPN profiles when the config.cmd deployment is made available through SCCM's Software Center. "Installing" the package performs the switch.

The script will always switch to Always ON on the first run, requiring the user to install the package twice the first time they want to switch to General. You may get around this by requiring a deployment running create_lockfile.cmd be installed first.

Returns error code 1 if jamCommand.exe is not found on the system.

Matt mcarras8