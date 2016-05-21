@echo off
RD /S /Q "commitaid"
DEL "*.exe"

md ".\commitaid"
md ".\commitaid\tools"

COPY ..\commit.vbs ".\commitaid\"
COPY ..\changelist.txt ".\commitaid\"
COPY ..\config.ini ".\commitaid\config_update.ini"

COPY ..\tools\install.vbs ".\commitaid\tools"

rar a -sfx ".\commitaid_2_0.exe" ".\commitaid"
rar c -zinstall_page_config.ini ".\commitaid_2_0.exe"
RD /S /Q "commitaid"