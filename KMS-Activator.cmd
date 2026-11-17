<# : KMS-Activator  (Batch + PowerShell hybrid script, double-click to run)
@echo off
setlocal
title KMS Activator
set "KMS_SELF=%~f0"
set "KMS_ARGS=%*"
fltmc >nul 2>&1 || (
    echo Requesting administrator privileges...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$a=@{FilePath=$env:KMS_SELF;Verb='RunAs'}; if($env:KMS_ARGS){$a.ArgumentList=$env:KMS_ARGS}; try{Start-Process @a}catch{Write-Host 'Administrator privileges are required.' -ForegroundColor Red; Start-Sleep 3; exit 1}"
    exit /b
)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$b=[IO.File]::ReadAllBytes($env:KMS_SELF);try{$s=(New-Object Text.UTF8Encoding($false,$true)).GetString($b)}catch{$s=[Text.Encoding]::Default.GetString($b)};iex $s.TrimStart([char]0xFEFF)"
exit /b %errorlevel%
#>

# =============================================================================
#  KMS 激活工具 -- 自动识别 Windows / Office 版本并匹配官方 GVLK
# -----------------------------------------------------------------------------
#  GVLK 来源 : Microsoft Learn《KMS 客户端激活和产品密钥》《Office 的 GVLK》
#  支持系统   : Windows 7 / 8 / 8.1 / 10 / 11，Server 2008 R2 ~ 2025
#  支持套件   : Office 2010 / 2013 / 2016 / 2019 / LTSC 2021 / LTSC 2024 (VL)
# =============================================================================

# ================================ 默认配置 ===================================
$Config = @{
    Server  = 'kms.cgtsoft.com'   # 默认 KMS 服务器（可通过菜单或 /server: 覆盖）
    Port    = 1688                # 默认 KMS 端口
    Timeout = 5000                # TCP 连通性测试超时（毫秒）
}
# =============================================================================

$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'KMS 激活工具 - Windows & Office' } catch { }

$WinAppId     = '55c92734-d682-4d71-983e-d6ec3f16059f'
$OfficeAppIds = @('0ff1ce15-a989-479d-af46-f275c6370663', '59a52881-a989-479d-af46-f275c6370663')

# =============================================================================
#  1. 数据字典：Windows & Office 官方 GVLK 与错误码说明
# =============================================================================
$GVLK = @{
    # Windows 10 / 11（常规渠道 & LTSC 2019 / 2021）
    'W10' = @{
        'Professional'             = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'
        'ProfessionalN'            = 'MH37W-N47XK-V7XM9-C7227-GCQG9'
        'ProfessionalWorkstation'  = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'
        'ProfessionalWorkstationN' = '9FNHH-K3HBT-3W4TD-6383H-6XYWF'
        'ProfessionalEducation'    = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'
        'ProfessionalEducationN'   = 'YVWGF-BXNMC-HTQYQ-CPQ99-66QFC'
        'Education'                = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'
        'EducationN'               = '2WH4N-8QGBV-H22JP-CT43Q-MDWWJ'
        'Enterprise'               = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'
        'EnterpriseN'              = 'DPH2V-TTNVB-4X9Q3-TJR4H-KHJW4'
        'EnterpriseG'              = 'YYVX9-NTFWV-6MDM3-9PT4T-4M68B'
        'EnterpriseGN'             = '44RPN-FTY23-9VTTB-MP9BX-T84FV'
        'ServerRdsh'               = 'CPWHC-NT2C7-VYW78-DHDB2-PG3GK'   # 企业版多会话
        'EnterpriseS'              = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'   # LTSC 2019 / 2021 / 2024
        'EnterpriseSN'             = '92NFX-8DJQP-P6BBQ-THF9C-7CG2H'
        'IoTEnterpriseS'           = 'KBN8V-HFGQ4-MGXVD-347P6-PDQGT'   # IoT Enterprise LTSC 2021
    }
    'W11LTSC2024' = @{
        'EnterpriseS'              = 'M7XTQ-FN8P6-TTKYV-9D4CC-J462D'
        'EnterpriseSN'             = '92NFX-8DJQP-P6BBQ-THF9C-7CG2H'
        'IoTEnterpriseS'           = 'CGK42-GYN6Y-VD22B-BX98W-J8JXD'   # IoT Enterprise LTSC 2024
    }
    'W10LTSB2016' = @{
        'EnterpriseS'              = 'DCPHK-NFMTC-H88MJ-PFHPY-QJ4BJ'
        'EnterpriseSN'             = 'QFFDN-GRT3P-VKWWX-X7T3R-8B639'
    }
    'W10LTSB2015' = @{
        'EnterpriseS'              = 'WNMTR-4C88C-JK8YV-HQ7T2-76DF9'
        'EnterpriseSN'             = '2F77B-TNFGY-69QQF-B8YKP-D69TJ'
    }
    'W81' = @{
        'Professional'             = 'GCRJD-8NW9H-F2CDX-CCM8D-9D6T9'
        'ProfessionalN'            = 'HMCNV-VVBFX-7HMBH-CTY9B-B4FXY'
        'Enterprise'               = 'MHF9N-XY6XB-WVXMC-BTDCT-MKKG7'
        'EnterpriseN'              = 'TT4HM-HN7YT-62K67-RGRQJ-JFFXW'
    }
    'W8' = @{
        'Professional'             = 'NG4HW-VH26C-733KW-K6F98-J8CK4'
        'ProfessionalN'            = 'XCVCF-2NXM9-723PB-MHCB7-2RYQQ'
        'Enterprise'               = '32JNW-9KQ84-P47T8-D8GGY-CWCK7'
        'EnterpriseN'              = 'JMNMF-RHW7P-DMY6X-RF3DR-X2BQT'
    }
    'W7' = @{
        'Professional'             = 'FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4'
        'ProfessionalN'            = 'MRPKT-YTG23-K7D7T-X2JMM-QY7MG'
        'ProfessionalE'            = 'W82YF-2Q76Y-63HXB-FGJG9-GF7QX'
        'Enterprise'               = '33PXH-7Y6KF-2VJC9-XBBR8-HVTHH'
        'EnterpriseN'              = 'YDRBP-3D83W-TY26F-D46B2-XCKRJ'
        'EnterpriseE'              = 'C29WB-22CC8-VJ326-GHFJW-H9DH4'
    }
    'S2025' = @{
        'ServerStandard'           = 'TVRH6-WHNXV-R9WG3-9XRFY-MY832'
        'ServerDatacenter'         = 'D764K-2NDRG-47T6Q-P8T8W-YP6DF'
        'ServerTurbine'            = 'XGN3F-F394H-FD2MY-PP6FD-8MCRC'   # Datacenter: Azure Edition
    }
    'S2022' = @{
        'ServerStandard'           = 'VDYBN-27WPP-V4HQT-9VMD4-VMK7H'
        'ServerDatacenter'         = 'WX4NM-KYWYW-QJJR4-XV3QB-6VM33'
        'ServerTurbine'            = 'NTBV8-9K7Q8-V27C6-M2BTV-KHMXV'   # Datacenter: Azure Edition
    }
    'S2019' = @{
        'ServerStandard'           = 'N69G4-B89J2-4G8F4-WWYCC-J464C'
        'ServerDatacenter'         = 'WMDGN-G9PQG-XVVXX-R3X43-63DFG'
        'ServerSolution'           = 'WVDHN-86M7X-466P6-VHXV7-YY726'   # Essentials
    }
    'S2016' = @{
        'ServerStandard'           = 'WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY'
        'ServerDatacenter'         = 'CB7KF-BWN84-R7R2Y-793K2-8XDDG'
        'ServerSolution'           = 'JCKRF-N37P4-C2D82-9YXRT-4M63B'   # Essentials
    }
    'S2012R2' = @{
        'ServerStandard'           = 'D2N9P-3P6X9-2R39C-7RTCD-MDVJX'
        'ServerDatacenter'         = 'W3GGN-FT8W3-Y4M27-J84CP-Q3VJ9'
        'ServerSolution'           = 'KNC87-3J2TX-XB4WP-VCPJV-M4FWM'   # Essentials
    }
    'S2012' = @{
        'ServerStandard'           = 'XC9B7-NBPP2-83J2H-RHMBY-92BT4'
        'ServerDatacenter'         = '48HP8-DN98B-MYWDG-T2DCC-8W83P'
        'ServerMultiPointStandard' = 'HM7DN-YVMH3-46JC3-XYTG7-CYQJJ'
        'ServerMultiPointPremium'  = 'XNH6W-2V9GX-RGJ4K-Y8X6F-QGJ2G'
    }
    'S2008R2' = @{
        'ServerWeb'                = '6TPJF-RBVHG-WBW2R-86QPH-6RTM4'
        'ServerHPC'                = 'TT8MH-CG224-D3D7Q-498W2-9QCTX'
        'ServerStandard'           = 'YC6KT-GKW9T-YTKYR-T4X34-R7VHC'
        'ServerEnterprise'         = '489J6-VHDMP-X63PK-3K798-CPX3Y'
        'ServerDatacenter'         = '74YFP-3QFB3-KQT8W-PMXWJ-7M648'
        'ServerEnterpriseIA64'     = 'GT63C-RJFQ3-4GMB6-BRFB9-CB83V'
    }
    'SAC1709' = @{
        'ServerDatacenterACor'     = '6Y6KB-N82V8-D8CQV-23MJW-BWTG6'
        'ServerStandardACor'       = 'DPCNP-XQFKJ-BJF7R-FRC8D-GF6G4'
    }
    'SAC1803' = @{
        'ServerDatacenterACor'     = '2HXDN-KRXHB-GPYC7-YCKFJ-7FVDG'
        'ServerStandardACor'       = 'PTXN8-JFHJM-4WC78-MPCBR-9W4KR'
    }
    'SAC' = @{
        'ServerDatacenterACor'     = '6NMRW-2C8FM-D24W7-TQWMY-CWH2D'
        'ServerStandardACor'       = 'N2KJX-J94YW-TQVFB-DG9YT-724CC'
    }
}

$FamilyName = @{
    'W10'         = 'Windows 10 / 11'
    'W11LTSC2024' = 'Windows 11 LTSC 2024'
    'W10LTSB2016' = 'Windows 10 LTSB 2016'
    'W10LTSB2015' = 'Windows 10 LTSB 2015'
    'W81'         = 'Windows 8.1'
    'W8'          = 'Windows 8'
    'W7'          = 'Windows 7'
    'S2025'       = 'Windows Server 2025'
    'S2022'       = 'Windows Server 2022'
    'S2019'       = 'Windows Server 2019'
    'S2016'       = 'Windows Server 2016'
    'S2012R2'     = 'Windows Server 2012 R2'
    'S2012'       = 'Windows Server 2012'
    'S2008R2'     = 'Windows Server 2008 R2'
    'SAC1709'     = 'Windows Server 1709 (SAC)'
    'SAC1803'     = 'Windows Server 1803 (SAC)'
    'SAC'         = 'Windows Server 1809-20H2 (SAC)'
}

$OfficeGVLK = @{
    '2024' = @{
        'ProPlus'          = 'XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB'
        'Standard'         = 'V28N4-JG22K-W66P8-VTMGK-H6HGR'
        'ProjectPro'       = 'FQQ23-N4YCY-73HQ3-FM9WC-76HF4'
        'ProjectStd'       = 'PD3TT-NTHQQ-VC7CY-MFXK3-G87F8'
        'VisioPro'         = 'B7TN8-FJ8V3-7QYCP-HQPMV-YY89G'
        'VisioStd'         = 'JMMVY-XFNQC-KK4HK-9H7R3-WQQTV'
        'Access'           = '82FTR-NCHR7-W3944-MGRHM-JMCWD'
        'Excel'            = 'F4DYN-89BP2-WQTWJ-GR8YC-CKGJG'
        'Outlook'          = 'D2F8D-N3Q3B-J28PV-X27HD-RJWB9'
        'PowerPoint'       = 'CW94N-K6GJH-9CTXY-MG2VC-FYCWP'
        'Word'             = 'MQ84N-7VYDM-FXV7C-6K7CC-VFW9J'
    }
    '2021' = @{
        'ProPlus'          = 'FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH'
        'Standard'         = 'KDX7X-BNVR8-TXXGX-4Q7Y8-78VT3'
        'ProjectPro'       = 'FTNWT-C6WBT-8HMGF-K9PRX-QV9H8'
        'ProjectStd'       = 'J2JDC-NJCYY-9RGQ4-YXWMH-T3D4T'
        'VisioPro'         = 'KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4'
        'VisioStd'         = 'MJVNY-BYWPY-CWV6J-2RKRT-4M8QG'
        'Access'           = 'WM8YG-YNGDD-4JHDC-PG3F4-FC4T4'
        'Excel'            = 'NWG3X-87C9K-TC7YY-BC2G7-G6RVC'
        'Outlook'          = 'C9FM6-3N72F-HFJXB-TM3V9-T86R9'
        'PowerPoint'       = 'TY7XF-NFRBR-KJ44C-G83KF-GX27K'
        'Publisher'        = '2MW9D-N4BXM-9VBPG-Q7W6M-KFBGQ'
        'SkypeForBusiness' = 'HWCXN-K3WBT-WJBKY-R8BD9-XK29P'
        'Word'             = 'TN8H9-M34D3-Y64V9-TR72V-X79KV'
    }
    '2019' = @{
        'ProPlus'          = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP'
        'Standard'         = '6NWWJ-YQWMR-QKGCB-6TMB3-9D9HK'
        'ProjectPro'       = 'B4NPR-3FKK7-T2MBV-FRQ4W-PKD2B'
        'ProjectStd'       = 'C4F7P-NCP8C-6CQPT-MQHV9-JXD2M'
        'VisioPro'         = '9BGNQ-K37YR-RQHF2-38RQ3-7VCBB'
        'VisioStd'         = '7TQNQ-K3YQQ-3PFH7-CCPPM-X4VQ2'
        'Access'           = '9N9PT-27V4Y-VJ2PD-YXFMF-YTFQT'
        'Excel'            = 'TMJWT-YYNMB-3BKTF-644FC-RVXBD'
        'Outlook'          = '7HD7K-N4PVK-BHBCQ-YWQRW-XW4VK'
        'PowerPoint'       = 'RRNCX-C64HY-W2MM7-MCH9G-TJHMQ'
        'Publisher'        = 'G2KWX-3NW6P-PY93R-JXK2T-C9Y9V'
        'SkypeForBusiness' = 'NCJ33-JHBBY-HTK98-MYCV8-HMKHJ'
        'Word'             = 'PBX3G-NWMT6-Q7XBW-PYJGG-WXD33'
    }
    '2016' = @{
        'ProPlus'          = 'XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99'
        'Standard'         = 'JNRGM-WHDWX-FJJG3-K47QV-DRTFM'
        'ProjectPro'       = 'YG9NW-3K39V-2T3HJ-93F3Q-G83KT'
        'ProjectStd'       = 'GNFHQ-F6YQM-KQDGJ-327XX-KQBVC'
        'VisioPro'         = 'PD3PC-RHNGV-FXJ29-8JK7D-RJRJK'
        'VisioStd'         = '7WHWN-4T7MP-G96JF-G33KR-W8GF4'
        'Access'           = 'GNH9Y-D2J4T-FJHGG-QRVH7-QPFDW'
        'Excel'            = '9C2PK-NWTVB-JMPW8-BFT28-7FTBF'
        'OneNote'          = 'DR92N-9HTF2-97XKM-XW2WJ-XW3J6'
        'Outlook'          = 'R69KK-NTPKF-7M3Q4-QYBHW-6MT9B'
        'PowerPoint'       = 'J7MQP-HNJ4Y-WJ7YM-PFYGF-BY6C6'
        'Publisher'        = 'F47MM-N3XJP-TQXJ9-BP99D-8K837'
        'SkypeForBusiness' = '869NQ-FJ69K-466HW-QYCP2-DDBV6'
        'Word'             = 'WXY84-JN2Q9-RBCCQ-3Q3J3-3PFJ6'
    }
    '2013' = @{
        'ProPlus'          = 'YC7DK-G2NP3-2QQC3-J6H88-GVGXT'
        'Standard'         = 'KBKQT-2NMXY-JJWGP-M62JB-92CD4'
        'ProjectPro'       = 'FN8TT-7WMH6-2D4X9-M337T-2342K'
        'ProjectStd'       = '6NTH3-CW976-3G3Y2-JK3TX-8QHTT'
        'VisioPro'         = 'C2FG9-N6J68-H8BTJ-BW3QX-RM3B3'
        'VisioStd'         = 'J484Y-4NKBF-W2HMG-DBMJC-PGWR7'
        'Access'           = 'NG2JY-H4JBT-HQXYP-78QH9-4JM2D'
        'Excel'            = 'VGPNG-Y7HQW-9RHP7-TKPV3-BG7GB'
        'InfoPath'         = 'DKT8B-N7VXH-D963P-Q4PHY-F8894'
        'Lync'             = '2MG3G-3BNTT-3MFW9-KDQW3-TCK7R'
        'OneNote'          = 'TGN6P-8MMBC-37P2F-XHXXK-P34VW'
        'Outlook'          = 'QPN8Q-BJBTJ-334K3-93TGY-2PMBT'
        'PowerPoint'       = '4NT99-8RJFH-Q2VDH-KYG2C-4RD4F'
        'Publisher'        = 'PN2WF-29XG2-T9HJ7-JQPJR-FCXK4'
        'Word'             = '6Q7VD-NX8JD-WJ2VH-88V73-4GBJ7'
    }
    '2010' = @{
        'ProPlus'          = 'VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB'
        'Standard'         = 'V7QKV-4XVVR-XYV4D-F7DFM-8R6BM'
        'ProjectPro'       = 'YGX6F-PGV49-PGW3J-9BTGG-VHKC6'
        'ProjectStd'       = '4HP3K-88W3F-W2K3D-6677X-F9PGB'
        'VisioPrem'        = 'D9DWC-HPYVV-JGF4P-BTWQB-WX8BJ'
        'VisioPro'         = '7MCW8-VRQVK-G677T-PDJCM-Q8TCP'
        'VisioStd'         = '767HD-QGMWX-8QTDB-9G3R2-KHFGJ'
        'Access'           = 'V7Y44-9T38C-R2VJK-666HK-T7DDX'
        'Excel'            = 'H62QG-HXVKF-PP4HP-66KMR-CW9BM'
        'InfoPath'         = 'K96W8-67RPQ-62T9Y-J8FQJ-BT37T'
        'OneNote'          = 'Q4Y4M-RHWJM-PY37F-MTKWH-D3XHX'
        'Outlook'          = '7YDC2-CWM8M-RRTJC-8MDVC-X3DWQ'
        'PowerPoint'       = 'RC8FX-88JRY-3PF7C-X8P67-P4VTT'
        'Publisher'        = 'BFK7F-9MYHM-V68C7-DRQ66-83YTP'
        'Groove'           = 'QYYW6-QP4CB-MBV6G-HYMCJ-4T3J4'   # SharePoint Workspace
        'Word'             = 'HVHB3-C6FV7-KQX9W-YQG79-CRY7T'
    }
}

$OfficeYear = @{ '14' = '2010'; '15' = '2013'; '16' = '2016'; '19' = '2019'; '21' = '2021'; '24' = '2024' }

$OfficeEditionAlias = @{
    'professionalplus'    = 'ProPlus'
    'proplus'             = 'ProPlus'
    'standard'            = 'Standard'
    'projectpro'          = 'ProjectPro'
    'projectprofessional' = 'ProjectPro'
    'projectstd'          = 'ProjectStd'
    'projectstandard'     = 'ProjectStd'
    'visioprem'           = 'VisioPrem'
    'visiopremium'        = 'VisioPrem'
    'visiopro'            = 'VisioPro'
    'visioprofessional'   = 'VisioPro'
    'visiostd'            = 'VisioStd'
    'visiostandard'       = 'VisioStd'
    'skypeforbusiness'    = 'SkypeForBusiness'
    'lync'                = 'Lync'
    'infopath'            = 'InfoPath'
    'onenote'             = 'OneNote'
    'powerpoint'          = 'PowerPoint'
    'publisher'           = 'Publisher'
    'outlook'             = 'Outlook'
    'access'              = 'Access'
    'excel'               = 'Excel'
    'word'                = 'Word'
    'groove'              = 'Groove'
}

$ErrorText = @{
    'C004F074' = '无法连接 KMS 服务器（服务器离线、端口被拦截或 DNS 解析失败）'
    '8007232B' = 'DNS 名称不存在：未配置 KMS 服务器或域名无法解析'
    '8007232A' = 'DNS 服务器故障'
    'C004F038' = 'KMS 服务器激活计数不足（服务器侧未达阈值，请稍后重试）'
    'C004F039' = 'KMS 服务器激活计数不足（服务器侧未达阈值，请稍后重试）'
    'C004F06C' = '时间戳无效：本机系统时间与 KMS 服务器相差过大，请先同步时间'
    'C004F050' = '产品密钥无效'
    'C004F069' = '找不到对应 SKU：当前软件可能不是批量授权（VL）版本'
    'C004F035' = '该计算机无法使用批量授权密钥激活（Win7 需 VL 版或含有效 SLIC 表）'
    'C004F015' = '许可证不受信任：当前版本可能不是批量授权（VL）版本'
    'C004C003' = '产品密钥已被阻止'
    'C004F042' = '指定的 KMS 服务器不支持激活此产品'
    'C004F041' = 'KMS 服务器本身未激活'
    'C004E016' = '产品密钥不支持当前版本'
    '80070005' = '拒绝访问：请右键[以管理员身份运行]'
    '80041003' = 'WMI 拒绝访问：请右键[以管理员身份运行]'
    '8007000D' = 'KMS 服务器返回的数据无效'
    'C004F056' = 'KMS 激活失败'
}

$StatusText = @{
    0 = '未授权'
    1 = '已激活'
    2 = '初始宽限期'
    3 = '附加宽限期'
    4 = '非正版宽限期'
    5 = '通知模式（未激活）'
    6 = '扩展宽限期'
}

# =============================================================================
#  2. 界面排版与输出工具（支持中日韩双宽字符精确对齐）
# =============================================================================
$script:LineWidth = 66

function Get-DisplayWidth([string]$Text) {
    if (-not $Text) { return 0 }
    $w = 0
    foreach ($ch in $Text.ToCharArray()) {
        if ([int]$ch -gt 0x7F) { $w += 2 } else { $w += 1 }
    }
    return $w
}

function Pad-RightDisplay([string]$Text, [int]$Width) {
    $w = Get-DisplayWidth $Text
    if ($w -ge $Width) { return $Text + ' ' }
    return $Text + (' ' * ($Width - $w))
}

function Center-Display([string]$Text, [int]$Width) {
    $w = Get-DisplayWidth $Text
    if ($w -ge $Width) { return $Text }
    $left = [int][math]::Floor(($Width - $w) / 2)
    $right = $Width - $w - $left
    return (' ' * $left) + $Text + (' ' * $right)
}

function Write-Line([string]$Text = '', [string]$Color = 'Gray') {
    Write-Host $Text -ForegroundColor $Color
}

function Write-Rule([string]$Char = '-', [string]$Color = 'DarkGray') {
    Write-Host ('  ' + ($Char * $script:LineWidth)) -ForegroundColor $Color
}

function Write-Section([string]$Title) {
    $label = ' ' + $Title + ' '
    $w = Get-DisplayWidth $label
    $left = 4
    $right = [math]::Max(4, $script:LineWidth - $left - $w)
    Write-Host ''
    Write-Host ('  ' + ('=' * $left)) -ForegroundColor DarkCyan -NoNewline
    Write-Host $label -ForegroundColor Cyan -NoNewline
    Write-Host ('=' * $right) -ForegroundColor DarkCyan
}

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host '  [*] ' -ForegroundColor Cyan -NoNewline
    Write-Host $Text -ForegroundColor White
}

function Write-Ok([string]$Text) {
    Write-Host '      [OK] ' -ForegroundColor Green -NoNewline
    Write-Host $Text -ForegroundColor Green
}

function Write-Warn([string]$Text) {
    Write-Host '      [!!] ' -ForegroundColor Yellow -NoNewline
    Write-Host $Text -ForegroundColor Yellow
}

function Write-Err([string]$Text) {
    Write-Host '      [XX] ' -ForegroundColor Red -NoNewline
    Write-Host $Text -ForegroundColor Red
}

function Write-Note([string]$Text) {
    Write-Host ('           ' + $Text) -ForegroundColor DarkYellow
}

function Write-KV([string]$Key, $Value, [string]$ValColor = 'White', [int]$KeyWidth = 14) {
    $padded = Pad-RightDisplay $Key $KeyWidth
    Write-Host ('      ' + $padded) -ForegroundColor Gray -NoNewline
    Write-Host ': ' -ForegroundColor DarkGray -NoNewline
    Write-Host $Value -ForegroundColor $ValColor
}

function Show-Banner {
    $border = '=' * $script:LineWidth
    $title1 = Center-Display 'KMS 激活工具  (Windows & Office)' ($script:LineWidth - 2)
    Write-Host ''
    Write-Host ('  ' + $border) -ForegroundColor DarkCyan
    Write-Host '  |' -ForegroundColor DarkCyan -NoNewline
    Write-Host $title1 -ForegroundColor Cyan -NoNewline
    Write-Host '|' -ForegroundColor DarkCyan
    Write-Host ('  ' + $border) -ForegroundColor DarkCyan
}

function Show-Help {
    Write-Host ''
    Write-Host '  用法: KMS-Activator.cmd [目标] [参数]' -ForegroundColor Cyan
    Write-Rule '-'
    Write-KV '(无参数)'            '打开交互菜单，选择激活 Windows 或 Office' 'White' 22
    Write-KV '/windows 或 /win'    '自动识别并激活 Windows'                   'White' 22
    Write-KV '/office'             '自动识别并激活 Office'                    'White' 22
    Write-KV '/all'                '依次自动激活 Windows 和 Office'           'White' 22
    Write-KV '/check'              '仅检测环境与匹配的 GVLK，不做任何修改'    'White' 22
    Write-KV '/status'             '查看当前 Windows 与 Office 激活状态'      'White' 22
    Write-KV '/server:主机[:端口]' '指定 KMS 服务器（默认 kms.cgtsoft.com）'  'White' 22
    Write-KV '/silent'             '静默运行：不询问、结束不暂停（默认 /all）' 'White' 22
    Write-KV '/force'              '即使 Windows 已永久激活也强制替换为 KMS'  'White' 22
    Write-KV '/help 或 /?'         '显示本帮助说明'                           'White' 22
    Write-Rule '-'
    Write-Host '  示例: KMS-Activator.cmd /office /silent' -ForegroundColor DarkGray
    Write-Host '        KMS-Activator.cmd /all /server:kms.cgtsoft.com:1688' -ForegroundColor DarkGray
    Write-Host ''
}

# =============================================================================
#  3. 参数解析
# =============================================================================
function Parse-ServerSpec([string]$Spec, $Opts) {
    if (-not $Spec) { return $false }
    $s = $Spec.Trim()
    if ($s -match '^\[(.+)\]:(\d+)$' -or $s -match '^([^:]+):(\d+)$') {
        $Opts.Server = $matches[1].Trim()
        $Opts.Port   = [int]$matches[2]
        return $true
    }
    $Opts.Server = $s.Trim('[', ']')
    return [bool]$Opts.Server
}

function Get-Options {
    $o = @{
        Silent  = $false; Force  = $false; Check  = $false; Status = $false; Help = $false
        Windows = $false; Office = $false
        Server  = $Config.Server; Port = [int]$Config.Port
    }
    $raw = $env:KMS_ARGS
    if (-not $raw) { return $o }
    foreach ($tok in ($raw -split '\s+')) {
        if (-not $tok) { continue }
        $t = $tok.Trim('"').TrimStart([char[]]'/-')
        $parts = @($t -split '[:=]', 2)
        $name = $parts[0].ToLower()
        $val = $null
        if ($parts.Count -gt 1) { $val = $parts[1] }
        switch ($name) {
            'silent'  { $o.Silent  = $true }
            's'       { $o.Silent  = $true }
            'force'   { $o.Force   = $true }
            'f'       { $o.Force   = $true }
            'check'   { $o.Check   = $true }
            'status'  { $o.Status  = $true }
            'help'    { $o.Help    = $true }
            'h'       { $o.Help    = $true }
            '?'       { $o.Help    = $true }
            'windows' { $o.Windows = $true }
            'win'     { $o.Windows = $true }
            'office'  { $o.Office  = $true }
            'all'     { $o.Windows = $true; $o.Office = $true }
            'server'  { [void](Parse-ServerSpec $val $o) }
            default   { Write-Warn ('忽略未知参数: ' + $tok) }
        }
    }
    return $o
}

# =============================================================================
#  4. 系统识别与通用 WMI / 网络工具
# =============================================================================
$script:OSInfoCache = $null
function Get-OSInfo {
    if ($script:OSInfoCache) { return $script:OSInfoCache }
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $build    = [int]$cv.CurrentBuildNumber
    $edition  = [string]$cv.EditionID
    $instType = [string]$cv.InstallationType

    if ($instType) { $isServer = $instType -like 'Server*' }
    else           { $isServer = $edition  -like 'Server*' }

    $caption = $null
    try { $caption = (Get-WmiObject -Class Win32_OperatingSystem).Caption } catch { }
    if (-not $caption) { $caption = [string]$cv.ProductName }
    # Windows 11 的注册表 ProductName 仍为 "Windows 10"，按内部版本号 >= 22000 修正
    if (-not $isServer -and $build -ge 22000) {
        $caption = $caption -replace 'Windows 10', 'Windows 11'
    }

    $ver = [string]$cv.DisplayVersion
    if (-not $ver) { $ver = [string]$cv.ReleaseId }
    if (-not $ver) { $ver = [string]$cv.CSDVersion }
    $ubr = ''
    if ($cv.UBR -ne $null) { $ubr = '.' + $cv.UBR }

    $script:OSInfoCache = New-Object PSObject -Property @{
        Caption  = $caption.Trim()
        Build    = $build
        BuildStr = ([string]$build + $ubr)
        Version  = $ver
        Edition  = $edition
        InstType = $instType
        IsServer = $isServer
        Arch     = $env:PROCESSOR_ARCHITECTURE
    }
    return $script:OSInfoCache
}

function Get-Family([int]$Build, [bool]$IsServer, [string]$Edition) {
    if ($IsServer) {
        if ($Edition -like '*ACor') {
            if ($Build -eq 16299) { return 'SAC1709' }
            if ($Build -eq 17134) { return 'SAC1803' }
            if ($Build -ge 17763) { return 'SAC' }
            return $null
        }
        if ($Build -ge 26000) { return 'S2025' }
        if ($Build -ge 20348) { return 'S2022' }
        if ($Build -ge 17763) { return 'S2019' }
        if ($Build -ge 14393) { return 'S2016' }
        if ($Build -ge 9600)  { return 'S2012R2' }
        if ($Build -ge 9200)  { return 'S2012' }
        if ($Build -ge 7600)  { return 'S2008R2' }
        return $null
    }
    if ($Build -ge 10240) {
        if ($Edition -eq 'EnterpriseS' -or $Edition -eq 'EnterpriseSN' -or $Edition -eq 'IoTEnterpriseS') {
            if ($Build -eq 10240) { return 'W10LTSB2015' }
            if ($Build -eq 14393) { return 'W10LTSB2016' }
            if ($Build -ge 26000) { return 'W11LTSC2024' }
        }
        return 'W10'
    }
    if ($Build -ge 9600) { return 'W81' }
    if ($Build -ge 9200) { return 'W8' }
    if ($Build -ge 7600) { return 'W7' }
    return $null
}

function Get-WinLicense {
    $filter = "ApplicationID='$WinAppId' AND PartialProductKey <> null"
    $list = @(Get-WmiObject -Class SoftwareLicensingProduct -Filter $filter | Where-Object { -not $_.LicenseIsAddon })
    if ($list.Count -gt 0) { return $list[0] }
    return $null
}

function Get-HResult($ErrRecord) {
    $e = $ErrRecord.Exception
    while ($e) {
        if ($e -is [System.Runtime.InteropServices.COMException]) { return [int]$e.ErrorCode }
        if ($e -is [System.Management.ManagementException]) { return [int]$e.ErrorCode }
        $e = $e.InnerException
    }
    return $null
}

function Format-Error($ErrRecord) {
    $hr = Get-HResult $ErrRecord
    if ($hr -eq $null) { return $ErrRecord.Exception.Message }
    $code = '{0:X8}' -f $hr
    $desc = $ErrorText[$code]
    if (-not $desc) { $desc = $ErrRecord.Exception.Message }
    return ('0x{0} - {1}' -f $code, $desc)
}

function Get-LicenseSummary($Lic) {
    if (-not $Lic) { return @{ Text = '未安装密钥'; Color = 'DarkYellow' } }
    $st = [int]$Lic.LicenseStatus
    $mins = [long]$Lic.GracePeriodRemaining
    if ($st -eq 1) {
        if ($mins -eq 0) { return @{ Text = '已永久激活'; Color = 'Green' } }
        $days = [int][math]::Floor($mins / 1440)
        return @{ Text = ('已激活 (剩余 {0} 天)' -f $days); Color = 'Green' }
    }
    $label = $StatusText[$st]
    if (-not $label) { $label = '状态码 ' + $st }
    if ($st -eq 2 -or $st -eq 3 -or $st -eq 6) {
        return @{ Text = $label; Color = 'Yellow' }
    }
    return @{ Text = $label; Color = 'Red' }
}

function Show-License($Lic) {
    if (-not $Lic) {
        Write-Warn '未找到已安装的产品密钥'
        return
    }
    $sum = Get-LicenseSummary $Lic
    Write-KV '产品名称' $Lic.Name
    Write-KV '授权渠道' $Lic.Description
    Write-KV '部分密钥' ('*****-' + $Lic.PartialProductKey) 'Yellow'
    Write-KV '激活状态' $sum.Text $sum.Color

    $mins = [long]$Lic.GracePeriodRemaining
    if ([int]$Lic.LicenseStatus -eq 1 -and $mins -eq 0) {
        Write-KV '有效期限' '永久有效' 'Green'
    } elseif ($mins -gt 0) {
        $exp = (Get-Date).AddMinutes($mins)
        $days = [math]::Floor($mins / 1440)
        Write-KV '有效期限' ('剩余 {0} 天（至 {1}）' -f $days, $exp.ToString('yyyy-MM-dd HH:mm'))
    }
    if ($Lic.Description -match 'KMSCLIENT') {
        $kms = $Lic.KeyManagementServiceMachine
        if (-not $kms) { $kms = $Lic.DiscoveredKeyManagementServiceMachineName }
        if ($kms) {
            $port = $Lic.KeyManagementServiceMachinePort
            if (-not $port) { $port = 1688 }
            Write-KV 'KMS 主机' ('{0}:{1}' -f $kms, $port) 'Cyan'
        }
    }
}

function Test-KmsPort([string]$HostName, [int]$Port) {
    $client = New-Object System.Net.Sockets.TcpClient
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $ar = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $ar.AsyncWaitHandle.WaitOne([int]$Config.Timeout, $false)) {
            return @{ Ok = $false; Ms = 0 }
        }
        $client.EndConnect($ar)
        $sw.Stop()
        return @{ Ok = $true; Ms = [int]$sw.ElapsedMilliseconds }
    } catch {
        return @{ Ok = $false; Ms = 0 }
    } finally {
        $client.Close()
    }
}

$script:KmsChecked = $null
function Assert-KmsServer($o) {
    if ($script:KmsChecked -ne $null) { return $script:KmsChecked }
    Write-Step ('测试 KMS 服务器连通性 ({0}:{1})' -f $o.Server, $o.Port)
    $ping = Test-KmsPort $o.Server $o.Port
    if ($ping.Ok) {
        Write-Ok ('连接成功（响应耗时 {0} ms）' -f $ping.Ms)
        $script:KmsChecked = $true
    } else {
        Write-Err '无法连接 KMS 服务器，请检查网络、防火墙，或更换 KMS 服务器。'
        $script:KmsChecked = $false
    }
    return $script:KmsChecked
}

function Confirm-Action([string]$Prompt) {
    Write-Host ''
    Write-Host ('      [?] ' + $Prompt + ' (y/N): ') -ForegroundColor Yellow -NoNewline
    $ans = Read-Host
    return ($ans -match '^(y|yes|是)$')
}

function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return (New-Object Security.Principal.WindowsPrincipal $id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# =============================================================================
#  5. Windows 激活模块
# =============================================================================
function Invoke-Windows($o) {
    Write-Section 'Windows 激活'

    # 1. 检测系统信息
    Write-Step '检测当前系统信息'
    $os = Get-OSInfo
    Write-KV '操作系统' $os.Caption 'Cyan'
    $verText = $os.BuildStr
    if ($os.Version) { $verText = '{0} (内部版本 {1})' -f $os.Version, $os.BuildStr }
    Write-KV '系统版本' $verText
    Write-KV '版本标识' $os.Edition
    Write-KV '系统架构' $os.Arch

    if ($o.Status) {
        Write-Step 'Windows 当前激活状态'
        Show-License (Get-WinLicense)
        return 0
    }

    # 2. 匹配 GVLK
    Write-Step '匹配微软官方 KMS 客户端密钥 (GVLK)'
    if ($os.Edition -like '*Eval*') {
        Write-Err '当前系统为评估版 (Evaluation)，不支持直接通过 KMS 激活。'
        Write-Note '请先使用 DISM /Online /Set-Edition 转换为正式版后再运行本工具。'
        return 3
    }
    $family = Get-Family $os.Build $os.IsServer $os.Edition
    $key = $null
    if ($family -and $GVLK.ContainsKey($family)) { $key = $GVLK[$family][$os.Edition] }
    if (-not $key) {
        Write-Err ('未找到与版本标识 "{0}" 匹配的 GVLK。' -f $os.Edition)
        Write-Note 'KMS 仅适用于专业版/企业版/教育版/工作站版/LTSC/Server；'
        Write-Note '家庭版 (Core / Home) 等零售版本无官方 KMS 密钥。'
        return 2
    }
    Write-KV '产品系列' $FamilyName[$family]
    Write-KV '匹配密钥' $key 'Yellow'
    Write-KV 'KMS 主机' ('{0}:{1}' -f $o.Server, $o.Port) 'Cyan'

    if ($o.Check) {
        Write-Ok 'Windows 环境检测与密钥匹配完成（仅检测模式，未修改系统）'
        return 0
    }

    if (-not (Test-Admin)) {
        Write-Err '需要管理员权限才能激活，请右键选择[以管理员身份运行]。'
        return 1
    }

    # 3. 检查现有激活状态，保护永久激活
    Write-Step '检查当前授权状态'
    $lic = Get-WinLicense
    Show-License $lic
    if ($lic -and [int]$lic.LicenseStatus -eq 1 -and [long]$lic.GracePeriodRemaining -eq 0 -and $lic.Description -notmatch 'KMSCLIENT') {
        Write-Warn '当前系统已处于永久激活状态（数字许可证 / 零售 / OEM / MAK）。'
        if (-not $o.Force) {
            if ($o.Silent) {
                Write-Note '静默模式下已自动跳过（如需强制替换请加 /force）。'
                return 0
            }
            if (-not (Confirm-Action '继续操作将替换为 180 天周期的 KMS 激活，是否继续？')) {
                Write-Warn '已取消 Windows 激活操作。'
                return 5
            }
        }
    }

    # 4. 测试 KMS 服务器连通性
    if (-not (Assert-KmsServer $o)) { return 4 }

    $svc = Get-WmiObject -Class SoftwareLicensingService

    # 5. 安装 GVLK
    Write-Step '安装 GVLK 产品密钥'
    $keyTail = $key.Substring($key.Length - 5)
    if ($lic -and $lic.PartialProductKey -eq $keyTail) {
        Write-Ok ('系统中已安装相同 GVLK (*****-{0})，无需重复安装' -f $keyTail)
    } else {
        try {
            [void]$svc.InstallProductKey($key)
            [void]$svc.RefreshLicenseStatus()
            Write-Ok ('成功安装密钥: ' + $key)
        } catch {
            Write-Err ('安装密钥失败: ' + (Format-Error $_))
            return 1
        }
    }

    # 6. 配置 KMS 并激活
    Write-Step '写入 KMS 服务器配置并请求激活'
    try {
        [void]$svc.SetKeyManagementServiceMachine($o.Server)
        [void]$svc.SetKeyManagementServicePort($o.Port)
        $lic = Get-WinLicense
        if (-not $lic) { throw '安装密钥后未能读取到 Windows 授权对象' }
        try { [void]$lic.ClearKeyManagementServiceMachine() } catch { }
        try { [void]$lic.ClearKeyManagementServicePort() } catch { }
        [void]$lic.Activate()
        [void]$svc.RefreshLicenseStatus()
        Write-Ok 'KMS 激活请求已成功完成'
    } catch {
        Write-Err ('激活失败: ' + (Format-Error $_))
        return 1
    }

    # 7. 验证结果
    Write-Step '验证 Windows 激活结果'
    $lic = Get-WinLicense
    Show-License $lic
    if ($lic -and [int]$lic.LicenseStatus -eq 1) {
        Write-Host ''
        Write-Ok 'Windows 激活成功！系统将在联网时自动向 KMS 服务器续期。'
        return 0
    }
    Write-Err 'Windows 激活未成功完成，请检查输出信息或稍后重试。'
    return 1
}

# =============================================================================
#  6. Office 激活模块
# =============================================================================
function Get-OfficeLicInfo($Wmi, [string]$ServiceClass) {
    $name = [string]$Wmi.Name
    $fam = ''
    if ($name -match ',\s*(\S+)\s+edition') { $fam = $matches[1] }
    elseif ($Wmi.LicenseFamily) { $fam = [string]$Wmi.LicenseFamily }

    $year = $null
    if ($fam -match '^Office(\d{2})') { $year = $OfficeYear[$matches[1]] }
    if (-not $year -and $fam -match '(20\d\d)') { $year = $matches[1] }
    if (-not $year -and $name -match 'Office\s*(\d{2})') { $year = $OfficeYear[$matches[1]] }

    $rawEd = ($fam -replace '^Office\d{0,2}', '') -replace '(20\d\d)?(VL)?[-_]?KMS_Client.*$', ''
    $canonEd = $OfficeEditionAlias[$rawEd.ToLower()]
    if (-not $canonEd) { $canonEd = $rawEd }

    $key = $null
    if ($year -and $canonEd -and $OfficeGVLK.ContainsKey($year)) {
        $key = $OfficeGVLK[$year][$canonEd]
    }

    $display = $name
    if ($year -and $canonEd) { $display = 'Office {0} {1}' -f $year, $canonEd }

    return New-Object PSObject -Property @{
        Wmi     = $Wmi
        Service = $ServiceClass
        Name    = $name
        Display = $display
        Year    = $year
        Edition = $canonEd
        Key     = $key
        IsKms   = ([string]$Wmi.Description -match 'KMSCLIENT')
        HasKey  = [bool]$Wmi.PartialProductKey
    }
}

function Get-OfficeLicenses {
    $sources = @(
        @{ Product = 'SoftwareLicensingProduct';        Service = 'SoftwareLicensingService' },
        @{ Product = 'OfficeSoftwareProtectionProduct'; Service = 'OfficeSoftwareProtectionService' }
    )
    $filter = ($OfficeAppIds | ForEach-Object { "ApplicationID='$_'" }) -join ' OR '
    $result = @()
    foreach ($s in $sources) {
        $items = @()
        try { $items = @(Get-WmiObject -Class $s.Product -Filter $filter -ErrorAction Stop) } catch { }
        foreach ($w in $items) { $result += Get-OfficeLicInfo $w $s.Service }
    }
    return $result
}

function Get-OfficeInstallInfo {
    $c2rIds = @()
    $lines  = @()
    try {
        $c = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
        if ($c.ProductReleaseIds) {
            $c2rIds = @($c.ProductReleaseIds -split '[,\s]+' | Where-Object { $_ })
            $lines += ('Click-to-Run ({0} / {1}): {2}' -f $c.VersionToReport, $c.Platform, $c.ProductReleaseIds)
        }
    } catch { }
    foreach ($v in @('14.0', '15.0', '16.0')) {
        foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office')) {
            $p = '{0}\{1}\Common\InstallRoot' -f $root, $v
            if (Test-Path $p) {
                $path = $null
                try { $path = (Get-ItemProperty $p -ErrorAction Stop).Path } catch { }
                if ($path -and ($path -notmatch 'root\\')) {
                    $lines += ('MSI 安装版 (Office {0}): {1}' -f $v, $path)
                }
            }
        }
    }
    return New-Object PSObject -Property @{ Lines = $lines; ReleaseIds = $c2rIds }
}

function Select-OfficeToInstall($KmsLics, [string[]]$ReleaseIds) {
    $haveKeys = @($KmsLics | Where-Object { $_.HasKey })
    $missing  = @($KmsLics | Where-Object { -not $_.HasKey -and $_.Key })
    if ($missing.Count -eq 0) { return @() }

    # 若 C2R 注册表列出了实际安装的 Volume 产品（如 ProPlus2021Volume, VisioPro2021Volume），按实际安装补齐缺失密钥
    if ($ReleaseIds -and @($ReleaseIds).Count -gt 0) {
        $matched = @()
        foreach ($m in $missing) {
            foreach ($rid in $ReleaseIds) {
                if ($rid -like ('*' + $m.Edition + '*')) {
                    $already = @($haveKeys | Where-Object { $_.Year -eq $m.Year -and $_.Edition -eq $m.Edition })
                    if ($already.Count -eq 0) { $matched += $m }
                    break
                }
            }
        }
        if ($matched.Count -gt 0) { return $matched }
    }

    # 若尚未安装任何 Office KMS 密钥，则为所有可匹配的 KMS 许可证安装 GVLK
    if ($haveKeys.Count -eq 0) { return $missing }
    return @()
}

function Invoke-Office($o) {
    Write-Section 'Office 激活'

    # 1. 检测安装
    Write-Step '检测已安装的 Office 组件'
    $inst = Get-OfficeInstallInfo
    if (@($inst.Lines).Count -gt 0) {
        foreach ($line in $inst.Lines) { Write-KV '安装信息' $line 'Cyan' }
    } else {
        Write-KV '安装信息' '未在注册表中发现 Office 安装记录' 'DarkGray'
    }

    $all     = @(Get-OfficeLicenses)
    $kms     = @($all | Where-Object { $_.IsKms })
    $withKey = @($all | Where-Object { $_.HasKey })

    if ($o.Status) {
        Write-Step 'Office 当前激活状态'
        if ($withKey.Count -eq 0) {
            Write-Warn '未找到已安装产品密钥的 Office 组件'
            return 0
        }
        foreach ($l in $withKey) {
            Write-Host ''
            Write-Host ('      [' + $l.Display + ']') -ForegroundColor Cyan
            Show-License $l.Wmi
        }
        return 0
    }

    # 2. 匹配批量授权许可证
    Write-Step '匹配 Office 批量授权许可证与官方 GVLK'
    if ($kms.Count -eq 0) {
        if (@($inst.Lines).Count -eq 0 -and $all.Count -eq 0) {
            Write-Err '未在本机检测到任何 Office 组件。'
            return 2
        }
        Write-Err '已安装的 Office 不是批量授权（VL）版本，无法通过 KMS 激活。'
        foreach ($l in $withKey) {
            Write-Note ('已安装渠道: {0} ({1})' -f $l.Display, $l.Wmi.Description)
        }
        Write-Note '零售版 / Microsoft 365 请使用正版账号或零售密钥激活；'
        Write-Note '如需 KMS 激活，请用 Office Tool Plus 或 ODT 安装 VL 批量版。'
        return 2
    }

    $targets   = @($kms | Where-Object { $_.HasKey })
    $toInstall = @(Select-OfficeToInstall $kms $inst.ReleaseIds)

    foreach ($l in $targets) {
        Write-KV $l.Display ('已安装 GVLK (*****-{0})' -f $l.Wmi.PartialProductKey) 'Green' 24
    }
    foreach ($l in $toInstall) {
        Write-KV $l.Display ('待安装 GVLK: ' + $l.Key) 'Yellow' 24
    }
    if ($targets.Count -eq 0 -and $toInstall.Count -eq 0) {
        foreach ($l in $kms) {
            Write-KV $l.Display '未能从许可证名称匹配到对应 GVLK' 'Red' 24
        }
        Write-Err '检测到 Office 批量授权许可证，但未能匹配到对应的 GVLK。'
        return 2
    }
    Write-KV 'KMS 主机' ('{0}:{1}' -f $o.Server, $o.Port) 'Cyan' 24

    if ($o.Check) {
        Write-Ok 'Office 环境检测与密钥匹配完成（仅检测模式，未修改系统）'
        return 0
    }

    if (-not (Test-Admin)) {
        Write-Err '需要管理员权限才能激活，请右键选择[以管理员身份运行]。'
        return 1
    }

    # 3. 测试 KMS 服务器连通性
    if (-not (Assert-KmsServer $o)) { return 4 }

    # 4. 安装缺失的 GVLK
    if ($toInstall.Count -gt 0) {
        Write-Step '安装 Office GVLK 产品密钥'
        foreach ($l in $toInstall) {
            try {
                $svc = Get-WmiObject -Class $l.Service
                [void]$svc.InstallProductKey($l.Key)
                [void]$svc.RefreshLicenseStatus()
                Write-Ok ('{0}: 成功安装 {1}' -f $l.Display, $l.Key)
            } catch {
                Write-Err ('{0}: 安装密钥失败 ({1})' -f $l.Display, (Format-Error $_))
            }
        }
        $targets = @(Get-OfficeLicenses | Where-Object { $_.IsKms -and $_.HasKey })
        if ($targets.Count -eq 0) {
            Write-Err '未能成功安装任何可激活的 Office GVLK。'
            return 1
        }
    }

    # 5. 逐个配置 KMS 并激活
    Write-Step '写入 KMS 服务器配置并激活 Office'
    foreach ($l in $targets) {
        $p = $l.Wmi
        $svc = Get-WmiObject -Class $l.Service
        try {
            # 优先设置在单个产品对象上，避免改动 Windows 系统的全局 KMS 主机设置
            $productLevel = $true
            try {
                [void]$p.SetKeyManagementServiceMachine($o.Server)
                [void]$p.SetKeyManagementServicePort($o.Port)
            } catch { $productLevel = $false }
            if (-not $productLevel) {
                [void]$svc.SetKeyManagementServiceMachine($o.Server)
                [void]$svc.SetKeyManagementServicePort($o.Port)
            }
            [void]$p.Activate()
            Write-Ok ($l.Display + ': 激活请求成功')
        } catch {
            Write-Err ('{0}: {1}' -f $l.Display, (Format-Error $_))
        }
        try { [void]$svc.RefreshLicenseStatus() } catch { }
    }

    # 6. 验证结果
    Write-Step '验证 Office 激活结果'
    $after = @(Get-OfficeLicenses | Where-Object { $_.IsKms -and $_.HasKey })
    $okCount = 0
    foreach ($l in $after) {
        Write-Host ''
        Write-Host ('      [' + $l.Display + ']') -ForegroundColor Cyan
        Show-License $l.Wmi
        if ([int]$l.Wmi.LicenseStatus -eq 1) { $okCount++ }
    }
    Write-Host ''
    if ($okCount -gt 0 -and $okCount -eq $after.Count) {
        Write-Ok 'Office 激活成功！软件将在联网时自动向 KMS 服务器续期。'
        return 0
    }
    if ($okCount -gt 0) {
        Write-Warn ('部分 Office 组件激活成功（{0}/{1}）' -f $okCount, $after.Count)
        return 1
    }
    Write-Err 'Office 激活未完成，请检查输出信息或稍后重试。'
    return 1
}

# =============================================================================
#  7. 交互式菜单与主入口
# =============================================================================
function Invoke-Targets($o) {
    $script:KmsChecked = $null
    $codes = @()
    if ($o.Windows) {
        try { $codes += Invoke-Windows $o }
        catch { Write-Err ('Windows 处理异常: ' + (Format-Error $_)); $codes += 1 }
    }
    if ($o.Office) {
        try { $codes += Invoke-Office $o }
        catch { Write-Err ('Office 处理异常: ' + (Format-Error $_)); $codes += 1 }
    }
    $worst = 0
    foreach ($c in $codes) {
        if ([int]$c -ne 0 -and [int]$c -ne 5) { $worst = [int]$c }
    }
    return $worst
}

function Wait-Key([string]$PromptText) {
    Write-Host ''
    Write-Rule '-'
    Write-Host ('  ' + $PromptText) -ForegroundColor DarkGray
    try { [void]$Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') } catch { [void](Read-Host) }
}

function Show-EnvironmentBar($o) {
    $os = Get-OSInfo
    $verStr = $os.BuildStr
    if ($os.Version) { $verStr = '{0} ({1})' -f $os.Version, $os.BuildStr }
    $adminText  = if (Test-Admin) { '管理员' } else { '普通用户 (激活需管理员权限)' }
    $adminColor = if (Test-Admin) { 'Green'  } else { 'Yellow' }

    Write-Host ''
    Write-KV '当前系统' ('{0} [{1}]' -f $os.Caption, $verStr) 'White' 12
    Write-KV '运行权限' $adminText $adminColor 12
    Write-KV 'KMS 主机' ('{0}:{1}' -f $o.Server, $o.Port) 'Cyan' 12
    Write-Rule '-'
}

function Write-MenuItem([string]$Key, [string]$Title, [string]$Desc = '', [string]$KeyColor = 'Cyan') {
    Write-Host '      [' -ForegroundColor DarkGray -NoNewline
    Write-Host $Key -ForegroundColor $KeyColor -NoNewline
    Write-Host '] ' -ForegroundColor DarkGray -NoNewline
    $padded = Pad-RightDisplay $Title 24
    Write-Host $padded -ForegroundColor White -NoNewline
    if ($Desc) { Write-Host $Desc -ForegroundColor DarkGray }
    else       { Write-Host '' }
}

function Prompt-Server($o) {
    Write-Host ''
    Write-Host ('      当前 KMS 主机 : {0}:{1}' -f $o.Server, $o.Port) -ForegroundColor Cyan
    Write-Host '      输入新地址 (回车恢复默认 kms.cgtsoft.com): ' -ForegroundColor Gray -NoNewline
    $inp = (Read-Host).Trim()
    if (-not $inp) {
        $o.Server = $Config.Server
        $o.Port   = [int]$Config.Port
    } else {
        [void](Parse-ServerSpec $inp $o)
    }
    Write-Ok ('已设置 KMS 服务器为 {0}:{1}' -f $o.Server, $o.Port)
    Start-Sleep -Milliseconds 700
}

function Show-Menu($o) {
    while ($true) {
        try { Clear-Host } catch { }
        Show-Banner
        Show-EnvironmentBar $o
        Write-Host ''
        Write-MenuItem '1' '激活 Windows'          '自动识别系统版本并激活'
        Write-MenuItem '2' '激活 Office'           '自动识别已安装的 Office 批量版并激活'
        Write-MenuItem '3' '激活 Windows + Office' '依次自动激活系统和 Office'
        Write-Host ''
        Write-MenuItem '4' '查看激活状态'          '查看 Windows 与 Office 详细授权状态' 'Yellow'
        Write-MenuItem '5' '仅检测匹配结果'        '测试版本识别与 GVLK 匹配，不做任何修改' 'Yellow'
        Write-MenuItem '6' '切换 KMS 服务器'       '临时更换本次会话使用的 KMS 主机地址' 'Yellow'
        Write-Host ''
        Write-MenuItem '0' '退出程序'              '' 'DarkGray'
        Write-Host ''
        Write-Rule '-'
        Write-Host '  请输入选项编号 [0-6]: ' -ForegroundColor Cyan -NoNewline
        $choice = (Read-Host).Trim()

        $o.Windows = $false; $o.Office = $false; $o.Check = $false; $o.Status = $false
        $runTarget = $true
        switch ($choice) {
            '1' { $o.Windows = $true }
            '2' { $o.Office  = $true }
            '3' { $o.Windows = $true; $o.Office = $true }
            '4' { $o.Windows = $true; $o.Office = $true; $o.Status = $true }
            '5' { $o.Windows = $true; $o.Office = $true; $o.Check  = $true }
            '6' { Prompt-Server $o; $runTarget = $false }
            '0' { return 0 }
            default { $runTarget = $false }
        }
        if (-not $runTarget) { continue }
        [void](Invoke-Targets $o)
        Wait-Key '按任意键返回主菜单...'
    }
}

$opts = Get-Options
$code = 0
try {
    if ($opts.Help) {
        Show-Banner
        Show-Help
        if (-not $opts.Silent) { Wait-Key '按任意键退出...' }
    } elseif ($opts.Windows -or $opts.Office -or $opts.Check -or $opts.Status -or $opts.Silent) {
        if (-not $opts.Windows -and -not $opts.Office) {
            $opts.Windows = $true
            $opts.Office  = $true
        }
        Show-Banner
        $code = Invoke-Targets $opts
        if (-not $opts.Silent) { Wait-Key '按任意键退出...' }
    } else {
        $code = Show-Menu $opts
    }
} catch {
    Write-Err ('发生未处理的错误: ' + (Format-Error $_))
    $code = 1
    if (-not $opts.Silent) { Wait-Key '按任意键退出...' }
}
exit $code
