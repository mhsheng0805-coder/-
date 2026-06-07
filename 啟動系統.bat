@echo off
chcp 65001 > nul
echo =====================================================
echo  紡織所 114年 來自民間業務收支管理系統
echo =====================================================
echo.
echo 正在啟動 Web 服務...
cd /d "%~dp0webapp"

:: 取得本機 IP
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /i "IPv4"') do (
    set LOCAL_IP=%%a
    goto :found
)
:found
set LOCAL_IP=%LOCAL_IP: =%

echo.
echo ✅ 系統已啟動！
echo.
echo  本機網址: http://127.0.0.1:5001
echo  區域網路: http://%LOCAL_IP%:5001
echo.
echo  在有 WiFi 的環境下，其他電腦/手機只要連上相同網路，
echo  在瀏覽器輸入上方「區域網路」網址即可填寫資料。
echo.
echo  關閉此視窗即停止服務。
echo =====================================================
start http://127.0.0.1:5001
python app.py
pause
