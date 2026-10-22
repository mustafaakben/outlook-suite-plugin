# Resolve all To/CC/BCC recipients before saving or sending a message.
function Assert-OutlookRecipients {
    param($Mail)
    if ($Mail.Recipients.Count -eq 0) { throw 'No recipients specified.' }
    if (-not $Mail.Recipients.ResolveAll()) {
        $unresolved = @(foreach ($recipient in $Mail.Recipients) {
            if (-not $recipient.Resolved) { $recipient.Name }
        })
        throw "Unresolved recipients: $($unresolved -join '; '). No message was saved or sent."
    }
}
