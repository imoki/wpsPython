<div align="center">
    <img src="https://socialify.git.ci/imoki/wpsPython/image?description=1&font=Rokkitt&forks=1&issues=1&language=1&owner=1&pattern=Circuit%20Board&pulls=1&stargazers=1&theme=Dark">
<h1>默定时代理</h1>
基于「金山文档」的python定时任务代理

<div id="shield">

[![][github-stars-shield]][github-stars-link]
[![][github-forks-shield]][github-forks-link]
[![][github-issues-shield]][github-issues-link]
[![][github-contributors-shield]][github-contributors-link]

<!-- SHIELD GROUP -->
</div>
</div>

## 👑 背景
金山文档目前不支持python脚本设置定时任务，导致许多python开发者转向不熟悉的airscript开发  
而本程序消除了金山文档的airscript调用python脚本的壁垒，使开发者能直接开发python脚本，并利用airscript的定时任务来运行python脚本。  
从而达到python也能定时执行的效果！  

## 🎊 简介
此脚本能够定时执行金山文档中的python脚本    
脚本能设置只调整哪些python脚本，能统一对多个文档内的多个python脚本进行个性化修改  
脚本具备多种灵活功能，在任意文档内的python脚本都能统一规划    

## ✨ 特性
    - 📀 支持金山文档运行
    - 💿 支持普通表格和智能表格
    - ♾️ 不限值python脚本位置
    - 💽 支持多文档统一修改
    

## 🍨 教程说明
💬 公众号“默库”

## 🛰️ 文字步骤
1. 将PY_INIT、PY脚本添加到金山文档中
2. 给PY_INIT、PY脚本添加网络API
3. 第一次运行PY_INIT脚本
4. 填写自动生成的wps表中的wps_sid
5. 再次运行PY_INIT脚本
6. 填写自动生成的PY表中的内容
7. 将PY脚本加入定时任务

## ⭐ 表格参考例子
![wps表](https://s3.bmp.ovh/imgs/2024/07/14/9045db168c0875ee.png)

## 🧾 表格内容含义 
1. wps_sid ： 填写wps文档内抓包得到的wps_sid
2. 文档名 : 填写需要修改定时任务时间的文档名称
3. 是否执行 ： 选项填“是”则会对其进行执行，默认为“否”是排除这个任务不会进行执行
4. 排除文档 ： 代表哪些文档不读取。以&分隔文档名，如：文档1&文档2
5. 仅读取文档 ： 代表仅读取哪些文档。以&分隔文档名，如：文档1&文档2。默认为@all代表所有文档都读取


## 🚀 其他
如果手动修改了定时任务时间，请重新运行一次CRON_INIT脚本，会自动生成最新的CRON配置表

## 🤝 欢迎参与贡献
欢迎各种形式的贡献

[![][pr-welcome-shield]][pr-welcome-link]

### 💗 感谢我们的贡献者
[![][github-contrib-shield]][github-contrib-link]


## ✨ Star 数

[![][starchart-shield]][starchart-link]

## 📝 更新日志 
- 2024-07-14
    * 推出金山文档python定时代理脚本

## 📌 特别声明

- 本仓库发布的脚本仅用于测试和学习研究，禁止用于商业用途，不能保证其合法性，准确性，完整性和有效性，请根据情况自行判断。

- 本人对任何脚本问题概不负责，包括但不限于由任何脚本错误导致的任何损失或损害。

- 间接使用脚本的任何用户，包括但不限于建立VPS或在某些行为违反国家/地区法律或相关法规的情况下进行传播, 本人对于由此引起的任何隐私泄漏或其他后果概不负责。

- 请勿将本仓库的任何内容用于商业或非法目的，否则后果自负。

- 如果任何单位或个人认为该项目的脚本可能涉嫌侵犯其权利，则应及时通知并提供身份证明，所有权证明，我们将在收到认证文件后删除相关脚本。

- 任何以任何方式查看此项目的人或直接或间接使用该项目的任何脚本的使用者都应仔细阅读此声明。本人保留随时更改或补充此免责声明的权利。一旦使用并复制了任何相关脚本或Script项目的规则，则视为您已接受此免责声明。

**您必须在下载后的24小时内从计算机或手机中完全删除以上内容**

> ***您使用或者复制了本仓库且本人制作的任何脚本，则视为 `已接受` 此声明，请仔细阅读***

<!-- LINK GROUP -->
[github-codespace-link]: https://codespaces.new/imoki/wpsPython
[github-codespace-shield]: https://github.com/imoki/wpsPython/blob/main/images/codespaces.png?raw=true
[github-contributors-link]: https://github.com/imoki/wpsPython/graphs/contributors
[github-contributors-shield]: https://img.shields.io/github/contributors/imoki/wpsPython?color=c4f042&labelColor=black&style=flat-square
[github-forks-link]: https://github.com/imoki/wpsPython/network/members
[github-forks-shield]: https://img.shields.io/github/forks/imoki/wpsPython?color=8ae8ff&labelColor=black&style=flat-square
[github-issues-link]: https://github.com/imoki/wpsPython/issues
[github-issues-shield]: https://img.shields.io/github/issues/imoki/wpsPython?color=ff80eb&labelColor=black&style=flat-square
[github-stars-link]: https://github.com/imoki/wpsPython/stargazers
[github-stars-shield]: https://img.shields.io/github/stars/imoki/wpsPython?color=ffcb47&labelColor=black&style=flat-square
[github-releases-link]: https://github.com/imoki/wpsPython/releases
[github-releases-shield]: https://img.shields.io/github/v/release/imoki/wpsPython?labelColor=black&style=flat-square
[github-release-date-link]: https://github.com/imoki/wpsPython/releases
[github-release-date-shield]: https://img.shields.io/github/release-date/imoki/wpsPython?labelColor=black&style=flat-square
[pr-welcome-link]: https://github.com/imoki/wpsPython/pulls
[pr-welcome-shield]: https://img.shields.io/badge/🤯_pr_welcome-%E2%86%92-ffcb47?labelColor=black&style=for-the-badge
[github-contrib-link]: https://github.com/imoki/wpsPython/graphs/contributors
[github-contrib-shield]: https://contrib.rocks/image?repo=imoki%2Fsign_script
[docker-pull-shield]: https://img.shields.io/docker/pulls/imoki/wpsPython?labelColor=black&style=flat-square
[docker-pull-link]: https://hub.docker.com/repository/docker/imoki/wpsPython
[docker-size-shield]: https://img.shields.io/docker/image-size/imoki/wpsPython?labelColor=black&style=flat-square
[docker-size-link]: https://hub.docker.com/repository/docker/imoki/wpsPython
[docker-stars-shield]: https://img.shields.io/docker/stars/imoki/wpsPython?labelColor=black&style=flat-square
[docker-stars-link]: https://hub.docker.com/repository/docker/imoki/wpsPython
[starchart-shield]: https://api.star-history.com/svg?repos=imoki/wpsPython&type=Date
[starchart-link]: https://api.star-history.com/svg?repos=imoki/wpsPython&type=Date

