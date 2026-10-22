# Shared outlook-suite helpers (connection, warm host, output, confirm gate, dates).
# They live in the mail skill; this bridge lets every skill dot-source the same way:
#   . (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "..\..\outlook-mail\scripts\_common.ps1")
