# �����ӵ�������������������
$name = "��̫��"
# �������� SSID
$ssid = "testssid"
# ������������
$key = "password"

Write "Internet Sharing Enabler v1.0 by James Swineson"

# chcp 437
# netsh wlan show drivers
# Write "> ֹͣ�ȵ�"
# netsh wlan stop hostednetwork
# Write "> �����ȵ�"
# netsh wlan set hostednetwork mode=disallow
Write "> �����������硱
netsh wlan set hostednetwork mode=allow ssid=$ssid key=$key
netsh wlan start hostednetwork
#netsh wlan show hostednetwork

Write "> ���ù�������ͨ�� $name ���ӵ�������"
# Register the HNetCfg library (once)
regsvr32 /s hnetcfg.dll

# Create a NetSharingManager object
$m = New-Object -ComObject HNetCfg.HNetShare

# List connections
# $m.EnumEveryConnection |% { $m.NetConnectionProps.Invoke($_) }

# Find connection
$c = $m.EnumEveryConnection |? { $m.NetConnectionProps.Invoke($_).Name -eq $name }

# Get sharing configuration
$config = $m.INetSharingConfigurationForINetConnection.Invoke($c)

# See if sharing is enabled
# Write-Output $config.SharingEnabled

# See the role of connection in sharing
# 0 - public, 1 - private
# Only meaningful if SharingEnabled is True
# Write-Output $config.SharingType

# Enable sharing (0 - public, 1 - private)
$config.EnableSharing(0)

# Disable sharing
# $config.DisableSharing()

# Write "��ǰ������״̬��"$config.SharingEnabled
Write "���в������"
Write "======================= ϵͳ����ӿ�״̬ ======================="
netsh wlan show interfaces
Write "======================= ������������״̬ ======================="
netsh wlan show hostednetwork
pause