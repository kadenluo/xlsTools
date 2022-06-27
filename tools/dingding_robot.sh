#!/bin/bash

author=$1
excel_file=$2
err_msg=$(echo "$3" | sed 's/"/\\"/g')
user_map=$(cat << EOF
pei.luo 13670178893
EOF
)

isAtAll="false"
phone=$(echo ${user_map} | grep -o "^${author}\>.*" | awk '{print $2}')
if [ "$phone" == "" ];then
    phone="$author"
    isAtAll="true"
fi

msg=$(cat << EOF
{
    "at": {
        "atMobiles":["${phone}"],
        "atUserIds":[],
        "isAtAll": ${isAtAll}
    },
    "text": {
	"content": "=====转表失败=====\n【excel】：${excel_file} \n【错误内容】：${err_msg}\n@${phone}",
    },
    "msgtype":"text"
}
EOF
)

curl -X POST -H "Content-Type: application/json" \
    -d "${msg}" \
    https://oapi.dingtalk.com/robot/send?access_token=c75df1f1adf83af86306608d5389bd7238cff6f1e68a347681d500d6c34651b4
