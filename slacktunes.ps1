$token = "xoxp-YOUR-TOKEN-HERE"
$itunes = New-Object  -ComObject iTunes.Application

function iTunes-Loop {
  $itunes = New-Object  -ComObject iTunes.Application
  if ($iTunes.PlayerState -eq "0"){
    echo "Not Playing"
    $status = @{
      profile = @{
        status_text = ""
        status_emoji = ""
        status_expiration = 0
      }
    } | ConvertTo-Json
    $headers = @{Authorization = "Bearer $token"}
    Invoke-WebRequest -UseBasicParsing https://slack.com/api/users.profile.set -Headers $headers -ContentType "application/json; charset=utf-8" -Method POST -Body $status | Out-Null
  }
  else {
    $artist = $iTunes.CurrentTrack.Artist
    $song = $iTunes.CurrentTrack.Name
    $status = "profile={`"status_text`": `"Listening to: $artist - $song`",`"status_emoji`": `":musical_note:`",`"status_expiration`": 0}"
    echo "Now Playing: $artist - $song"
    $status = @{
      profile = @{
        status_text = "Now Playing: $artist - $song"
        status_emoji = ":musical_note:"
        status_expiration = 0
      }
    } | ConvertTo-Json
    $headers = @{Authorization = "Bearer $token"}
    Invoke-WebRequest -UseBasicParsing https://slack.com/api/users.profile.set -Headers $headers -ContentType "application/json; charset=utf-8" -Method POST -Body $status | Out-Null
  }
}

do {
  $iTunesProcess = get-process "iTunes" -ea SilentlyContinue
  if ($iTunesProcess){
    iTunes-Loop
  }
  Start-Sleep -Seconds 60
}
until (-not $iTunesProcess)
echo "Not Running"
