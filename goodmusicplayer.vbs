While true
Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")
oPlayer.URL = "https://tinyurl.com/3m5yt37b"
oPlayer.controls.play
While oPlayer.playState <> 1 ' 1 = Stopped
WScript.Sleep 100
Wend
oPlayer.close
Wend
