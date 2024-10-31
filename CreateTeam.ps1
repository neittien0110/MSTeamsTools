#-------------------------------------------------
# Mục tiêu: 
#    Nếu thiết lập biến  Prompt = $True thì
#    - Nhập tên team từ bàn phím 
#    - Nhập tên của người đồng sáng lập. Enter luon nếu bỏ qua
#      Xong. Hệ thống sẽ tạo Team dạng Class rỗng ko có sinh viên và chưa được Active. Bạn hãy Active trên giao diện bằng cách click chuột phải vào biểu tượng Team đó
#    Nếu thiết lập biến Prompt = $False thì
#    - Gán trực tiếp tên Team trong mã nguồn
#    - Gán trực tiếp tên của người đồng sáng lập. Comment lại dòng lệnh nếu muốn bỏ qua.
#    - Copy danh sách email sinh viên vào biến, mỗi bạn 1 dòng.
#      Xong. Hệ thống sẽ tạo Team dạng Class với đủ các sinh viên làm member
#
# Trong trường hợp bạn chỉ muốn thêm thành viên vào một team đã có sẵn hãy bỏ qua (Enter luôn) khi được hỏi vể tên của Team, hoặc comment luôn dòng khai bao trong code
#-------------------------------------------------
# Điều kiện hoạt động
#  bạn cần cài đặt thêm module MicrosoftTeams cho PowerSheel. Gõ 2 lệnh sau. Tham khảo https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install
#   Install-Module -Name PowerShellGet -Force -AllowClobber 
#   Install-Module -Name MicrosoftTeams -Force -AllowClobber

<#
.SYNOPSIS
    Tao MSTeam moi, hoac su dung lai Team cu. Dong thoi bo sung them thanh vien vao Team do.Tao MSTeam moi, hoac su dung lai Team cu. Dong thoi bo sung them thanh vien vao Team do.
.DESCRIPTION
    - TeamName  Ten cua Team. Neu co ki tu space, hay dat trong dau "". Neu tham so TeamLink duoc khai bao, TeamName se bi bo qua.
    - TeamLink  GroupID cua Team cu da ton tai. De xem GroupID cua Team cu, hay vao Team va dung chuc nang <Get link to team> de lay URL co dang https://teams.microsoft.com/l/team/...  
    - CoOwner   Bo sung them 1 thanh vien sang lap owner. Vi du hoa.lt241234567@sis.hust.edu.vn
    - UserFile  Duong dan toi file danh sach chua cac thanh vien. Email cua moi thanh vien tren mot dong. Vi du
                tien.nguyen123@hust.edu.vn
                hoa.le456@hust.edu.vn
    Huong dan chi tiet: https://neittien0110.github.io/msteamstools

.EXAMPLE
   > CreateTeam.exe 
   > CreateTeam.ps1

.EXAMPLE
   > CreateTeam.exe -TeamLink "https://teams.microsoft.com/l/team/19%3Ac413f762004341f3b69d9fd6bb28aa0b%40thread.tacv2/conversations?groupId=ccdb66b9-449d-446d-a2bd-7f21199fe859&tenantId=06f1b89f-07e8-464f-b408-ec1b45703f31"  -UserFile  .\KyThuatMayTinh01-K64.txt
   > .\CreateTeam.ps1 -TeamLink "https://teams.microsoft.com/l/team/19%3Ac413f762004341f3b69d9fd6bb28aa0b%40thread.tacv2/conversations?groupId=ccdb66b9-449d-446d-a2bd-7f21199fe859&tenantId=06f1b89f-07e8-464f-b408-ec1b45703f31"  -UserFile  .\KyThuatMayTinh02-K64.txt

.EXAMPLE
   > CreateTeam.exe -CoOwner "abc@hust.edu.vn"  -UserFile  .\KyThuatMayTinh03-K64.txt
   > .\CreateTeam.ps1 -CoOwner "abc@hust.edu.vn"  -UserFile  .\KyThuatMayTinh04-K64.txt

.EXAMPLE
   > CreateTeam.exe -ChannelName "Kỹ thuật Máy tính 06 - K67" -UserFile  .\KyThuatMayTinh06-K67.txt
   > .\CreateTeam.ps1 -ChannelName "Kỹ thuật Máy tính 06 - K67" -UserFile  .\KyThuatMayTinh06-K67.txt

.LINK 
   GitHub: https://neittien0110.github.io/msteamstools

#>
## Get commandline params
Param(
    ## Ten team
    [Parameter(Mandatory=$false)]
    [string]
    $TeamName,
   
    ## Link cua Team
    [Parameter(Mandatory=$false)]
    [string]
    $TeamLink,

    ## Owner bo sung, ngoai nguoi tao
    [Parameter(Mandatory=$false)]
    [string]
    $CoOwner,
   
    ## Đương dẫn tới file chứa danh sách sinh viên
    [Parameter(Mandatory=$false)]
    [string]
    $UserFile,    
    
    ## Tên của channel sẽ add toàn bộ sinh viên theo danh sách vào đó. Không khai báo ($null) tức là bỏ qua, gõ "" tức là sẽ nhập từ bàn phím
    [Parameter(Mandatory=$false)]
    [string]
    $ChannelName          
)
Write-Host "Xem huong dan CreateTeam.exe -?"

# Danh sách các Module bắt buộc phải cài đặt
$Dependancies = @(                  
'PowerShellGet'
'MicrosoftTeams'
)

# Đồng sáng lập. Để trống nếu không muốn bổ sung owner.
$MyCoOwner = $CoOwner

# Danh sách email của sinh viên, mỗi sinh viên một dòng, có hoặc không có kí tự, cũng được
# Đọc file CreateTeam.student.txt, hoặc gán luôn theo đinh dạng
# $MyStudents = @(                  
#         'An.NV240057P@sis.hust.edu.vn'
#         'Anh.BD240058P@sis.hust.edu.vn'
# )

$MyStudents = @() 

function InputBox($formTitle, $textTitle){
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = $formTitle
    $objForm.Size = New-Object System.Drawing.Size(500,410)
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    #$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objTextBox.Text;$objForm.Close()}})  #Bỏ qua xử ly Enter là submit luôn, vì hộp textbox là multiline
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(75,340)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$Script:userInput=$objTextBox.Text;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CANCELButton = New-Object System.Windows.Forms.Button
    $CANCELButton.Location = New-Object System.Drawing.Size(250,340)
    $CANCELButton.Size = New-Object System.Drawing.Size(75,23)
    $CANCELButton.Text = "CANCEL"
    $CANCELButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CANCELButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,10)
    $objLabel.Size = New-Object System.Drawing.Size(320,40)
    $objLabel.Text = $textTitle
    $objForm.Controls.Add($objLabel)

    $objTextBox = New-Object System.Windows.Forms.TextBox
    $objTextBox.Location = New-Object System.Drawing.Size(10,60)
    $objTextBox.Multiline = $True
    $objTextBox.Size = New-Object System.Drawing.Size(460,270)
    $objForm.Controls.Add($objTextBox)

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})

    [void] $objForm.ShowDialog()

    return $userInput
}


function Test-Administrator 
{  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    return (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}

# Kiểm tra gói Package quan trọng đã cài đặt chưa
Write-Host "Kiem tra cac Package bat buoc..."
$RunAvaiable=$True
for ($i=0; $i -lt $Dependancies.count; $i++) {
    $IsInstalled = Get-Package | oss | Select-String -Pattern $Dependancies[$i]
    Write-Host ($i+1)"."$Dependancies[$i]'.........' $(If ($IsInstalled) {"ok"} Else {"khong"})  
    If ($IsInstalled) {
        $Dependancies[$i] = ''
    }else{
        $RunAvaiable=$False
    }
}

if (-Not $RunAvaiable) {
    Write-Host "Thieu goi bat buoc. Chuong trinh se tu dong cai dat ngay bay gio...."
    $IsAdmin = Test-Administrator
    if ($IsAdmin -eq $False) {
        Write-Host "- Chuong trinh khong chay voi quyen Administrator, nen se cai dat cho duy nhat User hien tai."
    }
    for ($i=0; $i -lt $Dependancies.count; $i++) {
        If ($Dependancies[$i] -ne '') {
            Write-Host ($i+1)"."$Dependancies[$i]'......... be installing' 
            Install-Module -Name $Dependancies[$i] -Force -AllowClobber -Scope CurrentUser  # Giới hạn scope để không đòi quyền Admin
        }
    }
}

Write-Host ''

#------------------------ KẾT NỐI MICROSOFT TEAMS ---------------------

#Kết nôi Microsoft Teams. Vui lòng đăng nhập trên giao diên web khi được hỏi
Write-Host "Cua so dang nhap Microsoft Teams dang mo. Vui long nhap thong tin tai khoan..."
Connect-MicrosoftTeams 
if ( $? -eq $False ) {
    Write-Host " - Khong dang nhap duoc vao Teams. Ket thuc"
    exit -1
}

#------------------------ LÂY THÔNG TIN GROUP ID CỦA TEAM CŨ ---------------------

if ($TeamLink -eq "" -And $TeamName -eq "") {
    Write-Host "Ban muon them thanh vien vao Team cu, hay dung chuc nang <Get link to team> de lay URL co dang https://teams.microsoft.com/l/team/..."
    $TeamLink=Read-Host "Hay nhap URL Link to Team cua Team hien co (muon them thanh vien), hoac Enter de bo qua va tao Team moi:" 
}

#------------------------ TẠO TEAM MỚI  ---------------------

if ($TeamLink -eq "" )  # Nếu có tên Team đầy đủ thì mới tạo team
{
    if ($TeamName -eq "") {
        $TeamName = Read-Host "Nhap ten Team moi:"    
    }
    Write-Host "Tao Team "  $TeamName

    #Tạo Team ở dạng Class với Tên như chỉ định
    $group = New-Team  -template EDU_Class -DisplayName $TeamName
    # Kiểm tra lai qua trình tạo Team mới
    if ( $? -eq $False ) {
        Write-Host " - Khong tao duoc Team moi. Ket thuc"
        exit -1
    }    
    #Lấy lai GroupID da tao
    $MyTeamGroupID = $group.GroupId
} else {
    $MyTeamGroupID  = $($TeamLink.Split("`?=&"))[2] 
    $group = $(Get-Team -GroupID $MyTeamGroupID)
    if ( $? -eq $False ) {
        Write-Host " - Khong truy cap duoc URL da cho. Ket thuc"
        exit -1
    }        
    $TeamName = $group.Description
}
Write-Host " - Team name $TeamName"
Write-Host " - GroupId $MyTeamGroupID .Done"

#------------------------ BỔ SUNG OWNER ---------------------

if ($MyCoOwner -ne "") {
    #Bổ sung owner cho kênh mới (nếu có)
    Write-Host "Them co-owner "  $MyCoOwner
    Add-TeamUser -GroupId $MyTeamGroupID -Role Owner -User $MyCoOwner    
}

#------------------------ BỔ SUNG THÀNH VIÊN ---------------------
if ($UserFile -eq "") {
    Write-Host 'Nhap danh sach sinh vien trong hop InpuxBox'
    $tmp = InputBox("Danh sach thanh vien","Dang ky cac thanh vien vao team. Moi nguoi 1 dong. Vi du: an.nv202022@sis.hust.edu.vn binh.nt29393929@sis.hust.edu.vn")
    $MyStudents=$tmp.Split([Environment]::NewLine, [StringSplitOptions]::RemoveEmptyEntries)
}else{
    Write-Host "Kiem tra danh sach sinh vien trong file $UserFile"
    $MyStudents = Get-Content -Path $UserFile
}
Write-Host  " - Co"$MyStudents.count"thanh vien can dang ky."
if ($MyStudents.count -eq 0){
    exit(0);
}
Write-Host  " - 1."$MyStudents[0]
Write-Host  " - $($MyStudents.count). "$MyStudents[$($MyStudents.count-1)]

$tmp = Read-Host "Bam [Y] de bat dau them thanh vien:" # Tên Team
if (($tmp -ne "Y") -And ($tmp -ne "y")) {
    exit(0);
}

#Bổ sung sinh viên cho Teams mới (nếu có)
for ($i=0; $i -lt $MyStudents.count; $i++) {
    Write-Host ($i+1)". " $MyStudents[$i] 
    Add-TeamUser -GroupId $MyTeamGroupID -Role Member -User $MyStudents[$i] -ErrorAction SilentlyContinue
    if ( $? -eq $False ) {
        Write-Host " - Error: Khong them duoc user $($MyStudents[$i]) do khong ton tai."
        continue
    }         

}

#------------------------ BỔ SUNG CHANNEL MỚI ---------------------

# Nếu chưa khai báo thi hỏi thăm nhập liệu 
if ( $ChannelName -eq '' ) {
    Write-Host 'Bo sung cac thanh vien vao channel private'
    $ChannelName = Read-Host "Nhap ten Channel moi (Enter luon de bo qua):"
}

# Nếu đã có tên thì kiểm tra xem tên có đúng không
if ( $ChannelName -ne '' ) {
    # Lấy tất cả các handler của các channel
    $channels = Get-TeamChannel -GroupID $MyTeamGroupID 

    # Lọc theo tên channel 
    $exists = $channels | Where-Object { $_.DisplayName -eq $ChannelName }

    $MyChannel = $null
    # Lấy handler điều khiển channel. Tự động tạo channel nếu chưa tồn tại
    if ($exists) {
        Write-Output "Channel '$ChannelName' da co, san sang."
        $MyChannel = $exists[0]
    } else {
        Write-Output "Channel '$ChannelName' chua co. Tao moi"
        $MyChannel = New-TeamChannel -GroupId $MyTeamGroupID -MembershipType Private -DisplayName $ChannelName -Description "Auto gen, $(Get-Date)"
    }

    #Bổ sung sinh viên cho Channel mới (nếu có)
    for ($i=0; $i -lt $MyStudents.count; $i++) {
        Write-Host ($i+1)". " $MyStudents[$i] 
        Add-TeamChannelUser -GroupId $MyTeamGroupID -DisplayName $MyChannel.DisplayName -User $MyStudents[$i] -ErrorAction SilentlyContinue
        if ( $? -eq $False ) {
            Write-Host " - Error: Khong them duoc user $($MyStudents[$i]) vao channel $($MyChannel.DisplayName)."
            continue
        }     
    }
}


# Ngắt kết nôi với Teams
Disconnect-MicrosoftTeams

# groupId=ccdb66b9-449d-446d-a2bd-7f21199fe859&tenantId=06f1b89f-07e8-464f-b408-ec1b45703f31
#         7fcff8d3-f0ef-4c52-b097-288768349c88&tenantId=06f1b89f-07e8-464f-b408-ec1b45703f31