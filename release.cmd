set target=x86_64-pc-windows-msvc
xargo build --target %target% --release
upx --best --lzma .\target\%target%\release\volume_ctl.exe