@echo off
RD /S /Q "commitaid"
DEL "commitaid.exe"

md ".\commitaid"
md ".\commitaid\tools"

COPY ..\commit.vbs ".\commitaid\"
COPY ..\changelist.txt ".\commitaid\"
COPY ..\config.ini ".\commitaid\config_update.ini"

COPY ..\tools\install.vbs ".\commitaid\tools"

rar a -sfx ".\commitaid.exe" ".\commitaid"
rar c -zinstall_page_config.ini ".\commitaid.exe"
RD /S /Q "commitaid"