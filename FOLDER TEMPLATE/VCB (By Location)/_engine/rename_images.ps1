[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "   BAT 1: RENAME IMAGES (VCB BY LOCATION)          " -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
$root = Split-Path $PSScriptRoot -Parent
Write-Host "Working Directory: $root`n"

$folderMapping = @{
    "001. OVERVIEW FRONT (LHS)" = @{ Visual = "1.jpg"; IR = "2.jpg" }
    "002. OVERVIEW FRONT (RHS)" = @{ Visual = "3.jpg"; IR = "4.jpg" }
    "003. OVERVIEW REAR (LHS)" = @{ Visual = "5.jpg"; IR = "6.jpg" }
    "004. OVERVIEW REAR (RHS)" = @{ Visual = "7.jpg"; IR = "8.jpg" }
    "005. OVERVIEW TOP (LHS)" = @{ Visual = "9.jpg"; IR = "10.jpg" }
    "006. OVERVIEW TOP (RHS)" = @{ Visual = "11.jpg"; IR = "12.jpg" }
    "007. OVERVIEW BUSBAR (LHS)" = @{ Visual = "13.jpg"; IR = "14.jpg" }
    "008. OVERVIEW BUSBAR (RHS)" = @{ Visual = "15.jpg"; IR = "16.jpg" }
    "009. OVERVIEW BUSBAR (LHS)" = @{ Visual = "17.jpg"; IR = "18.jpg" }
    "010. OVERVIEW BUSBAR (RHS)" = @{ Visual = "19.jpg"; IR = "20.jpg" }
    "011. OVERVIEW CABLE CELLAR (LHS)" = @{ Visual = "21.jpg"; IR = "22.jpg" }
    "012. OVERVIEW CABLE CELLAR (RHS)" = @{ Visual = "23.jpg"; IR = "24.jpg" }
    "013. BREAKER COMPARTMENT - PANEL 1" = @{ Visual = "25.jpg"; IR = "26.jpg" }
    "014. BREAKER COMPARTMENT - PANEL 2" = @{ Visual = "41.jpg"; IR = "42.jpg" }
    "015. BREAKER COMPARTMENT - PANEL 3" = @{ Visual = "57.jpg"; IR = "58.jpg" }
    "016. BREAKER COMPARTMENT - PANEL 4" = @{ Visual = "73.jpg"; IR = "74.jpg" }
    "017. BREAKER COMPARTMENT - PANEL 5" = @{ Visual = "89.jpg"; IR = "90.jpg" }
    "018. BREAKER COMPARTMENT - PANEL 6" = @{ Visual = "105.jpg"; IR = "106.jpg" }
    "019. BREAKER COMPARTMENT - PANEL 7" = @{ Visual = "121.jpg"; IR = "122.jpg" }
    "020. BREAKER COMPARTMENT - PANEL 8" = @{ Visual = "137.jpg"; IR = "138.jpg" }
    "021. BREAKER COMPARTMENT - PANEL 9" = @{ Visual = "153.jpg"; IR = "154.jpg" }
    "022. BREAKER COMPARTMENT - PANEL 10" = @{ Visual = "169.jpg"; IR = "170.jpg" }
    "023. BREAKER COMPARTMENT - PANEL 11" = @{ Visual = "185.jpg"; IR = "186.jpg" }
    "024. BREAKER COMPARTMENT - PANEL 12" = @{ Visual = "201.jpg"; IR = "202.jpg" }
    "025. BREAKER COMPARTMENT - PANEL 13" = @{ Visual = "217.jpg"; IR = "218.jpg" }
    "026. BREAKER COMPARTMENT - PANEL 14" = @{ Visual = "233.jpg"; IR = "234.jpg" }
    "027. BREAKER COMPARTMENT - PANEL 15" = @{ Visual = "249.jpg"; IR = "250.jpg" }
    "028. BUSBAR DISCONNECTING SW - PANEL 1" = @{ Visual = "27.jpg"; IR = "28.jpg" }
    "029. BUSBAR DISCONNECTING SW - PANEL 2" = @{ Visual = "43.jpg"; IR = "44.jpg" }
    "030. BUSBAR DISCONNECTING SW - PANEL 3" = @{ Visual = "59.jpg"; IR = "60.jpg" }
    "031. BUSBAR DISCONNECTING SW - PANEL 4" = @{ Visual = "75.jpg"; IR = "76.jpg" }
    "032. BUSBAR DISCONNECTING SW - PANEL 5" = @{ Visual = "91.jpg"; IR = "92.jpg" }
    "033. BUSBAR DISCONNECTING SW - PANEL 6" = @{ Visual = "107.jpg"; IR = "108.jpg" }
    "034. BUSBAR DISCONNECTING SW - PANEL 7" = @{ Visual = "123.jpg"; IR = "124.jpg" }
    "035. BUSBAR DISCONNECTING SW - PANEL 8" = @{ Visual = "139.jpg"; IR = "140.jpg" }
    "036. BUSBAR DISCONNECTING SW - PANEL 9" = @{ Visual = "155.jpg"; IR = "156.jpg" }
    "037. BUSBAR DISCONNECTING SW - PANEL 10" = @{ Visual = "171.jpg"; IR = "172.jpg" }
    "038. BUSBAR DISCONNECTING SW - PANEL 11" = @{ Visual = "187.jpg"; IR = "188.jpg" }
    "039. BUSBAR DISCONNECTING SW - PANEL 12" = @{ Visual = "203.jpg"; IR = "204.jpg" }
    "040. BUSBAR DISCONNECTING SW - PANEL 13" = @{ Visual = "219.jpg"; IR = "220.jpg" }
    "041. BUSBAR DISCONNECTING SW - PANEL 14" = @{ Visual = "235.jpg"; IR = "236.jpg" }
    "042. BUSBAR DISCONNECTING SW - PANEL 15" = @{ Visual = "251.jpg"; IR = "252.jpg" }
    "043. CABLE COMPARTMENT - PANEL 1" = @{ Visual = "29.jpg"; IR = "30.jpg" }
    "044. CABLE COMPARTMENT - PANEL 2" = @{ Visual = "45.jpg"; IR = "46.jpg" }
    "045. CABLE COMPARTMENT - PANEL 3" = @{ Visual = "61.jpg"; IR = "62.jpg" }
    "046. CABLE COMPARTMENT - PANEL 4" = @{ Visual = "77.jpg"; IR = "78.jpg" }
    "047. CABLE COMPARTMENT - PANEL 5" = @{ Visual = "93.jpg"; IR = "94.jpg" }
    "048. CABLE COMPARTMENT - PANEL 6" = @{ Visual = "109.jpg"; IR = "110.jpg" }
    "049. CABLE COMPARTMENT - PANEL 7" = @{ Visual = "125.jpg"; IR = "126.jpg" }
    "050. CABLE COMPARTMENT - PANEL 8" = @{ Visual = "141.jpg"; IR = "142.jpg" }
    "051. CABLE COMPARTMENT - PANEL 9" = @{ Visual = "157.jpg"; IR = "158.jpg" }
    "052. CABLE COMPARTMENT - PANEL 10" = @{ Visual = "173.jpg"; IR = "174.jpg" }
    "053. CABLE COMPARTMENT - PANEL 11" = @{ Visual = "189.jpg"; IR = "190.jpg" }
    "054. CABLE COMPARTMENT - PANEL 12" = @{ Visual = "205.jpg"; IR = "206.jpg" }
    "055. CABLE COMPARTMENT - PANEL 13" = @{ Visual = "221.jpg"; IR = "222.jpg" }
    "056. CABLE COMPARTMENT - PANEL 14" = @{ Visual = "237.jpg"; IR = "238.jpg" }
    "057. CABLE COMPARTMENT - PANEL 15" = @{ Visual = "253.jpg"; IR = "254.jpg" }
    "058. BUSBAR COMPARTMENT - PANEL 1" = @{ Visual = "31.jpg"; IR = "32.jpg" }
    "059. BUSBAR COMPARTMENT - PANEL 2" = @{ Visual = "47.jpg"; IR = "48.jpg" }
    "060. BUSBAR COMPARTMENT - PANEL 3" = @{ Visual = "63.jpg"; IR = "64.jpg" }
    "061. BUSBAR COMPARTMENT - PANEL 4" = @{ Visual = "79.jpg"; IR = "80.jpg" }
    "062. BUSBAR COMPARTMENT - PANEL 5" = @{ Visual = "95.jpg"; IR = "96.jpg" }
    "063. BUSBAR COMPARTMENT - PANEL 6" = @{ Visual = "111.jpg"; IR = "112.jpg" }
    "064. BUSBAR COMPARTMENT - PANEL 7" = @{ Visual = "127.jpg"; IR = "128.jpg" }
    "065. BUSBAR COMPARTMENT - PANEL 8" = @{ Visual = "143.jpg"; IR = "144.jpg" }
    "066. BUSBAR COMPARTMENT - PANEL 9" = @{ Visual = "159.jpg"; IR = "160.jpg" }
    "067. BUSBAR COMPARTMENT - PANEL 10" = @{ Visual = "175.jpg"; IR = "176.jpg" }
    "068. BUSBAR COMPARTMENT - PANEL 11" = @{ Visual = "191.jpg"; IR = "192.jpg" }
    "069. BUSBAR COMPARTMENT - PANEL 12" = @{ Visual = "207.jpg"; IR = "208.jpg" }
    "070. BUSBAR COMPARTMENT - PANEL 13" = @{ Visual = "223.jpg"; IR = "224.jpg" }
    "071. BUSBAR COMPARTMENT - PANEL 14" = @{ Visual = "239.jpg"; IR = "240.jpg" }
    "072. BUSBAR COMPARTMENT - PANEL 15" = @{ Visual = "255.jpg"; IR = "256.jpg" }
    "073. PT COMPARTMENT - PANEL 1" = @{ Visual = "33.jpg"; IR = "34.jpg" }
    "074. PT COMPARTMENT - PANEL 2" = @{ Visual = "49.jpg"; IR = "50.jpg" }
    "075. PT COMPARTMENT - PANEL 3" = @{ Visual = "65.jpg"; IR = "66.jpg" }
    "076. PT COMPARTMENT - PANEL 4" = @{ Visual = "81.jpg"; IR = "82.jpg" }
    "077. PT COMPARTMENT - PANEL 5" = @{ Visual = "97.jpg"; IR = "98.jpg" }
    "078. PT COMPARTMENT - PANEL 6" = @{ Visual = "113.jpg"; IR = "114.jpg" }
    "079. PT COMPARTMENT - PANEL 7" = @{ Visual = "129.jpg"; IR = "130.jpg" }
    "080. PT COMPARTMENT - PANEL 8" = @{ Visual = "145.jpg"; IR = "146.jpg" }
    "081. PT COMPARTMENT - PANEL 9" = @{ Visual = "161.jpg"; IR = "162.jpg" }
    "082. PT COMPARTMENT - PANEL 10" = @{ Visual = "177.jpg"; IR = "178.jpg" }
    "083. PT COMPARTMENT - PANEL 11" = @{ Visual = "193.jpg"; IR = "194.jpg" }
    "084. PT COMPARTMENT - PANEL 12" = @{ Visual = "209.jpg"; IR = "210.jpg" }
    "085. PT COMPARTMENT - PANEL 13" = @{ Visual = "225.jpg"; IR = "226.jpg" }
    "086. PT COMPARTMENT - PANEL 14" = @{ Visual = "241.jpg"; IR = "242.jpg" }
    "087. PT COMPARTMENT - PANEL 15" = @{ Visual = "257.jpg"; IR = "258.jpg" }
    "088. CABLE ENTRY - PANEL 1" = @{ Visual = "35.jpg"; IR = "36.jpg" }
    "089. CABLE ENTRY - PANEL 2" = @{ Visual = "51.jpg"; IR = "52.jpg" }
    "090. CABLE ENTRY - PANEL 3" = @{ Visual = "67.jpg"; IR = "68.jpg" }
    "091. CABLE ENTRY - PANEL 4" = @{ Visual = "83.jpg"; IR = "84.jpg" }
    "092. CABLE ENTRY - PANEL 5" = @{ Visual = "99.jpg"; IR = "100.jpg" }
    "093. CABLE ENTRY - PANEL 6" = @{ Visual = "115.jpg"; IR = "116.jpg" }
    "094. CABLE ENTRY - PANEL 7" = @{ Visual = "131.jpg"; IR = "132.jpg" }
    "095. CABLE ENTRY - PANEL 8" = @{ Visual = "147.jpg"; IR = "148.jpg" }
    "096. CABLE ENTRY - PANEL 9" = @{ Visual = "163.jpg"; IR = "164.jpg" }
    "097. CABLE ENTRY - PANEL 10" = @{ Visual = "179.jpg"; IR = "180.jpg" }
    "098. CABLE ENTRY - PANEL 11" = @{ Visual = "195.jpg"; IR = "196.jpg" }
    "099. CABLE ENTRY - PANEL 12" = @{ Visual = "211.jpg"; IR = "212.jpg" }
    "100. CABLE ENTRY - PANEL 13" = @{ Visual = "227.jpg"; IR = "228.jpg" }
    "101. CABLE ENTRY - PANEL 14" = @{ Visual = "243.jpg"; IR = "244.jpg" }
    "102. CABLE ENTRY - PANEL 15" = @{ Visual = "259.jpg"; IR = "260.jpg" }
    "103. LINK BOX - PANEL 1" = @{ Visual = "37.jpg"; IR = "38.jpg" }
    "104. LINK BOX - PANEL 2" = @{ Visual = "53.jpg"; IR = "54.jpg" }
    "105. LINK BOX - PANEL 3" = @{ Visual = "69.jpg"; IR = "70.jpg" }
    "106. LINK BOX - PANEL 4" = @{ Visual = "85.jpg"; IR = "86.jpg" }
    "107. LINK BOX - PANEL 5" = @{ Visual = "101.jpg"; IR = "102.jpg" }
    "108. LINK BOX - PANEL 6" = @{ Visual = "117.jpg"; IR = "118.jpg" }
    "109. LINK BOX - PANEL 7" = @{ Visual = "133.jpg"; IR = "134.jpg" }
    "110. LINK BOX - PANEL 8" = @{ Visual = "149.jpg"; IR = "150.jpg" }
    "111. LINK BOX - PANEL 9" = @{ Visual = "165.jpg"; IR = "166.jpg" }
    "112. LINK BOX - PANEL 10" = @{ Visual = "181.jpg"; IR = "182.jpg" }
    "113. LINK BOX - PANEL 11" = @{ Visual = "197.jpg"; IR = "198.jpg" }
    "114. LINK BOX - PANEL 12" = @{ Visual = "213.jpg"; IR = "214.jpg" }
    "115. LINK BOX - PANEL 13" = @{ Visual = "229.jpg"; IR = "230.jpg" }
    "116. LINK BOX - PANEL 14" = @{ Visual = "245.jpg"; IR = "246.jpg" }
    "117. LINK BOX - PANEL 15" = @{ Visual = "261.jpg"; IR = "262.jpg" }
    "118. SECONDARY COMPARTMENT - PANEL 1" = @{ Visual = "39.jpg"; IR = "40.jpg" }
    "119. SECONDARY COMPARTMENT - PANEL 2" = @{ Visual = "55.jpg"; IR = "56.jpg" }
    "120. SECONDARY COMPARTMENT - PANEL 3" = @{ Visual = "71.jpg"; IR = "72.jpg" }
    "121. SECONDARY COMPARTMENT - PANEL 4" = @{ Visual = "87.jpg"; IR = "88.jpg" }
    "122. SECONDARY COMPARTMENT - PANEL 5" = @{ Visual = "103.jpg"; IR = "104.jpg" }
    "123. SECONDARY COMPARTMENT - PANEL 6" = @{ Visual = "119.jpg"; IR = "120.jpg" }
    "124. SECONDARY COMPARTMENT - PANEL 7" = @{ Visual = "135.jpg"; IR = "136.jpg" }
    "125. SECONDARY COMPARTMENT - PANEL 8" = @{ Visual = "151.jpg"; IR = "152.jpg" }
    "126. SECONDARY COMPARTMENT - PANEL 9" = @{ Visual = "167.jpg"; IR = "168.jpg" }
    "127. SECONDARY COMPARTMENT - PANEL 10" = @{ Visual = "183.jpg"; IR = "184.jpg" }
    "128. SECONDARY COMPARTMENT - PANEL 11" = @{ Visual = "199.jpg"; IR = "200.jpg" }
    "129. SECONDARY COMPARTMENT - PANEL 12" = @{ Visual = "215.jpg"; IR = "216.jpg" }
    "130. SECONDARY COMPARTMENT - PANEL 13" = @{ Visual = "231.jpg"; IR = "232.jpg" }
    "131. SECONDARY COMPARTMENT - PANEL 14" = @{ Visual = "247.jpg"; IR = "248.jpg" }
    "132. SECONDARY COMPARTMENT - PANEL 15" = @{ Visual = "263.jpg"; IR = "264.jpg" }
    "133. TX1 - OVERVIEW" = @{ Visual = "265.jpg"; IR = "266.jpg" }
    "134. TX1 - OVERVIEW TOP" = @{ Visual = "267.jpg"; IR = "268.jpg" }
    "135. TX1 - BUSHING (HV)" = @{ Visual = "269.jpg"; IR = "270.jpg" }
    "136. TX1 - CABLE (HV)" = @{ Visual = "271.jpg"; IR = "272.jpg" }
    "137. TX1 - CABLE SPLIT (HV)" = @{ Visual = "273.jpg"; IR = "274.jpg" }
    "138. TX1 - BUSHING (LV)" = @{ Visual = "275.jpg"; IR = "276.jpg" }
    "139. TX1 - CABLE (LV)" = @{ Visual = "277.jpg"; IR = "278.jpg" }
    "140. TX1 - CABLE SPLIT (LV)" = @{ Visual = "279.jpg"; IR = "280.jpg" }
    "141. TX2 - OVERVIEW" = @{ Visual = "281.jpg"; IR = "282.jpg" }
    "142. TX2 - OVERVIEW TOP" = @{ Visual = "283.jpg"; IR = "284.jpg" }
    "143. TX2 - BUSHING (HV)" = @{ Visual = "285.jpg"; IR = "286.jpg" }
    "144. TX2 - CABLE (HV)" = @{ Visual = "287.jpg"; IR = "288.jpg" }
    "145. TX2 - CABLE SPLIT (HV)" = @{ Visual = "289.jpg"; IR = "290.jpg" }
    "146. TX2 - BUSHING (LV)" = @{ Visual = "291.jpg"; IR = "292.jpg" }
    "147. TX2 - CABLE (LV)" = @{ Visual = "293.jpg"; IR = "294.jpg" }
    "148. TX2 - CABLE SPLIT (LV)" = @{ Visual = "295.jpg"; IR = "296.jpg" }
    "149. FP1 - OVERVIEW" = @{ Visual = "297.jpg"; IR = "298.jpg" }
    "150. FP2 - OVERVIEW" = @{ Visual = "299.jpg"; IR = "300.jpg" }
    "151. BLACK BOX 1 - OVERVIEW" = @{ Visual = "301.jpg"; IR = "302.jpg" }
    "152. BLACK BOX 2 - OVERVIEW" = @{ Visual = "303.jpg"; IR = "304.jpg" }
    "153. BATTERY BANK 1 - OVERVIEW" = @{ Visual = "305.jpg"; IR = "306.jpg" }
    "154. BATTERY BANK 2 - OVERVIEW" = @{ Visual = "307.jpg"; IR = "308.jpg" }
    "155. BATTERY BANK 3 - OVERVIEW" = @{ Visual = "309.jpg"; IR = "310.jpg" }
    "156. BATTERY BANK 4 - OVERVIEW" = @{ Visual = "311.jpg"; IR = "312.jpg" }
}

$subfolders = Get-ChildItem -Path $root -Directory | Where-Object { $_.Name -match '^\d+' }

if ($subfolders.Count -eq 0) {
    Write-Host "[ERROR] No numbered subfolders found in $root" -ForegroundColor Red
    return
}

$renamedCount = 0

foreach ($folder in ($subfolders | Sort-Object Name)) {
    if ($folderMapping.ContainsKey($folder.Name)) {
        $mapInfo = $folderMapping[$folder.Name]
        $visualImgName = $mapInfo.Visual
        $irImgName = $mapInfo.IR
        
        $files = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match '(?i)\.(jpg|jpeg|png)$' }
        if ($files.Count -eq 0) { continue }
        
        Write-Host "[$($folder.Name)]" -ForegroundColor Cyan
        
        $visualFile = $null
        $irFile = $null
        
        if ($files.Count -eq 1) {
            $f = $files[0]
            if ($f.Name -match '(?i)(photo|dc|vis|digital)') { $visualFile = $f } else { $irFile = $f }
        } elseif ($files.Count -ge 2) {
            foreach ($f in $files) {
                if ($f.Name -match '(?i)(photo|dc|vis|digital)') { $visualFile = $f }
                elseif ($f.Name -match '(?i)(flir|ir|thermal)') { $irFile = $f }
            }
            if (-not $visualFile -or -not $irFile) {
                $sorted = $files | Sort-Object Length -Descending
                $visualFile = $sorted[0]
                $irFile = $sorted[1]
            }
        }
        
        if ($visualFile -and $visualFile.Name -ne $visualImgName) {
            $targetPath = Join-Path $folder.FullName $visualImgName
            if (Test-Path $targetPath) { Remove-Item -Force $targetPath }
            Rename-Item -Path $visualFile.FullName -NewName $visualImgName -Force
            Write-Host "  [VISUAL] $($visualFile.Name) -> $visualImgName" -ForegroundColor Green
            $renamedCount++
        }
        
        if ($irFile -and $irFile.Name -ne $irImgName) {
            $targetPath = Join-Path $folder.FullName $irImgName
            if (Test-Path $targetPath) { Remove-Item -Force $targetPath }
            Rename-Item -Path $irFile.FullName -NewName $irImgName -Force
            Write-Host "  [IR]     $($irFile.Name) -> $irImgName" -ForegroundColor Green
            $renamedCount++
        }
    }
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host "  COMPLETED: $renamedCount images renamed in their folders." -ForegroundColor Yellow
Write-Host "===================================================" -ForegroundColor Cyan
