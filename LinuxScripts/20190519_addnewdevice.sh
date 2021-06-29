#! /bin/bash
while true
do

#clear
echo "============================================================="
echo "-------------------------------------------------------------"
echo "1. 备份配置文件（Device、device1、device2）"
echo "2. 测试rules文件语法"
echo "3. 重启进程"
echo "4. 同步配置到备机"
echo "5. 退出程序"
echo "-------------------------------------------------------------"

read -p "请选择一个选项（1-5）：" U_SELECT
case $U_SELECT in 
	1)
	echo "1. 备份配置文件：(1)Device, (2)device1, (3)device2..."
	read -p "If you want backup Device(y/n): " isbackupdevice
	if [ "$isbackupdevice" == "y" ]; then
		cp ./device/Device ./device/Device_20190508
		echo "(1)Device备份结果：" $(ls ./device/Device_20190508)
		echo
	fi
	
	read -p "是否备份device1文件（y/n）：" isbackupccbptwldevice
	if [ "$isbackupccbptwldevice" == "y" ]; then
		cp ./device/device1 ./device/device1_20190508
		echo "(2)device1备份结果：" $(ls ./device/device1_20190508)
		echo
	fi

	read -p "是否备份device2文件（y/n）：" isbackuphwdevice
	if [ "$isbackuphwdevice" == "y" ]; then
		cp ./device/device2 ./device/device2_20190508
		echo "(3)device2备份结果：" $(ls ./device/device2_20190508)
		echo
	fi
;;

	2)
	echo "2. 测试rules文件语法"
	echo
	echo
;;

3)
	echo "3. 重启进程"
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