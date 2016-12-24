# Windows Logon Scripts

## Overview
気が向いて作った Microsoft Windows グループポリシーログオンスクリプトの掃き溜め

## Contents

- GroupMembershipNetworkDriveMap
 - 所属グループに合わせ、指定ドライブへネットワークドライブをマウントするためのログオンスクリプト。
 - Launchapp.cmd をログオンスクリプトとして指定する。
 - 同一フォルダに GrpMemDrvMap.vbs、GrpMemDrvMapList.txt、[Launchapp.wsf](http://www.jhouseconsulting.com/2012/09/03/an-improved-and-enhanced-version-of-the-famous-launchapp-wsf-838) を配置する。
 - GrpMemDrvMapList.txt を編集して、所属グループ／マウント先のドライブレター／ネットワークドライブの場所／ドライブの名前を列挙する。
