#!/bin/bash
PATH=/bin:/sbin:/usr/bin:/usr/sbin:/usr/local/bin:/usr/local/sbin:~/bin
export PATH

#================================================================
#	System Required: macOS Sierra - macOS Monterey
#	Description: Modify MSOffice language on macOS
#	Version: 1.0
#	Author: Vincent Young
# 	Telegram: https://t.me/missuo
#	Github: https://github.com/missuo/MSOffice-Language-Modify
#	Latest Update: December 15, 2021
#=================================================================

# Change language to Chinese
change_language_to_chinese(){
    defaults write com.microsoft.Word AppleLanguages '("zh_CN")'
    defaults write com.microsoft.Excel AppleLanguages '("zh_CN")'
    defaults write com.microsoft.Powerpoint AppleLanguages '("zh_CN")'
    echo -e "Modified successfully! Current language is zh_CN."
    echo -e ""
}

# Change language to English
change_language_to_english(){
    defaults write com.microsoft.Word AppleLanguages '("en")'
    defaults write com.microsoft.Excel AppleLanguages '("en")'
    defaults write com.microsoft.Powerpoint AppleLanguages '("en")'
    echo -e "Modified successfully! Current language is en."
    echo -e ""
}

# change language custmize
change_language_customize(){
    read -p "Enter what language you want to change: " language
	[ -z "${language}" ]
	echo ""
    defaults write com.microsoft.Word AppleLanguages '("'${language}'")'
    defaults write com.microsoft.Excel AppleLanguages '("'${language}'")'
    defaults write com.microsoft.Powerpoint AppleLanguages '("'${language}'")'
    echo -e "Modified successfully! Current language is ${language}."
    echo -e ""
}

# Check current language
check_current_language(){
    echo -e "Current Word language is: "
    return_language=`defaults read com.microsoft.Word AppleLanguages`
    echo $return_language | cut -d '(' -f2 | cut -d ')' -f1
    echo -e ""
}

# Start Menu
start_menu(){
		clear
		echo && echo -e "Modify MSOffice language on macOS
Feedback： https://github.com/missuo/MSOffice-Language-Modify
————————————Mode Selection————————————
${green}1.${plain} Change language to Chinese
${green}2.${plain} Change language to English
${green}3.${plain} Change language to Customize
${green}4.${plain} Check current language
${green}0.${plain} Exit
——————————————————————————————————————"
	read -p "Please select: " num
	case "$num" in
	1)
		change_language_to_chinese
		;;
	2)
		change_language_to_english
		;;
    3)
        change_language_customize
        ;;
    4)
        check_current_language
        ;;
	0)
		exit 1
		;;
	*)
		clear
		echo -e "[${red}Error${plain}]:Please enter the correct number[0-4]."
		sleep 5s
		start_menu
		;;
	esac
}
start_menu 