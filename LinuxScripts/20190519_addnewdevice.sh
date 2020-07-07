#! /bin/bash
while true
do

#clear
echo "============================================================="
echo "-------------------------------------------------------------"
echo "1. 备份配置文件（Device、ccb_ptwl_device、HW_Devices）"
echo "2. 测试rules文件语法"
echo "3. 重启网管进程"
echo "4. 同步配置到备机"
echo "5. 退出程序"
echo "-------------------------------------------------------------"

read -p "请选择一个选项（1-5）：" U_SELECT
case $U_SELECT in 
	1)
	echo "1. 备份配置文件：(1)Device, (2)ccb_ptwl_device, (3)HW_Devices..."
	read -p "If you want backup Device(y/n): " isbackupdevice
	if [ "$isbackupdevice" == "y" ]; then
		cp /mnt/hgfs/ProjectsLinux/NMS/device/Device /mnt/hgfs/ProjectsLinux/NMS/device/Device_20190508
		echo "(1)Device备份结果：" $(ls /mnt/hgfs/ProjectsLinux/NMS/device/Device_20190508)
		echo
	fi
	
	read -p "是否备份ccb_ptwl_device文件（y/n）：" isbackupccbptwldevice
	if [ "$isbackupccbptwldevice" == "y" ]; then
		cp /mnt/hgfs/ProjectsLinux/NMS/device/ccb_ptwl_device /mnt/hgfs/ProjectsLinux/NMS/device/ccb_ptwl_device_20190508
		echo "(2)ccb_ptwl_device备份结果：" $(ls /mnt/hgfs/ProjectsLinux/NMS/device/ccb_ptwl_device_20190508)
		echo
	fi

	read -p "是否备份HW_Devices文件（y/n）：" isbackuphwdevice
	if [ "$isbackuphwdevice" == "y" ]; then
		cp /mnt/hgfs/ProjectsLinux/NMS/device/HW_Devices /mnt/hgfs/ProjectsLinux/NMS/device/HW_Devices_20190508
		echo "(3)HW_Devices备份结果：" $(ls /mnt/hgfs/ProjectsLinux/NMS/device/HW_Devices_20190508)
		echo
	fi
;;

	2)
	echo "2. 测试rules文件语法"
	echo
	echo
;;

3)
	echo "3. 重启网管进程"
	echo
	echo
;;

4)
	echo $(df -h)
	echo
	echo
;;

5)
	exit
;;

	*)
	read -p "请输入1-5的数字，按回车键继续："
	esac

done