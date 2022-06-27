#!/bin/bash
# 在crontab中添加以下配置即可：
# */1 * * * * export SVN_USER="xxx";export SVN_USER_PASSWORD="xxx";bash ./tools/autoForSvn.sh >> /tmp/cron.log 2>&1
# 使用前先修改下面的三个变量:xls_dir、client_dir和server_dir

cd $(dirname $0)/../
echo "=================($(date))======================="

xls_dir="../../CommonRes/Table/Trunk"
client_dir="../../Assets/Script/Lua/DataTables/CommonData"
server_dir="../../../Server/Server/DataTables/CommonData"

# 清理无关文件
svn cleanup ${xls_dir} --remove-unversioned
svn cleanup ${client_dir} --remove-unversioned
svn cleanup ${server_dir} --remove-unversioned

# try update
updateInfo="$(svn up ${xls_dir} --username ${SVN_USER} --password ${SVN_USER_PASSWORD})"

log_commit="$(svn log -l 1 ${xls_dir} --username ${SVN_USER} --password ${SVN_USER_PASSWORD}| awk '{if (NR==2){printf "%s %s ", $1,$3}else if(NR==4){print $1}}')"
log_version=$(echo ${log_commit} | awk '{print $1}')
log_author=$(echo ${log_commit} | awk '{print $2}')
log_message=$(echo ${log_commit} | awk '{print $3}')
commit_msg="${log_version}-${log_author}-${log_message}:转表"

echo "last-commit:${commit_msg}"

# 开始导表
res=$(bash ./xlsTools.sh --force)
if [ $? -ne 0 ];then
	echo "notify auther:$res"
	bash ./tools/dingding_robot.sh ${SVN_USER} "unknown" "${res}"
	exit -1
fi

function svn_commit()
{
	dir=$1
	changes=$(svn status ${dir})
	if [ "$changes" == "" ];then
		return
	fi
	svn add ${dir} 2> /dev/null
	svn commit ${dir} -m ${commit_msg} --username ${SVN_USER} --password ${SVN_USER_PASSWORD}
	if [ $? -ne 0 ];then
		echo "Error:submit failed."
		bash ./tools/dingding_robot.sh ${SVN_USER} "unknown" "svn提交失败"
		return
	fi
}

svn_commit ${client_dir}
svn_commit ${server_dir}
