<div align="center">

# 🗺️ AddressMap Tool
### 批量地址坐标转换 · 一个顶三个在线工具

[![Version](https://img.shields.io/badge/version-v2.8.66-blue.svg)](https://github.com/Mine1178/addressmap-tool/releases)
[![License](https://img.shields.io/badge/license-Apache--2.0-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![Python](https://img.shields.io/badge/Python-3.8+-yellow.svg)]()

**通信勘察 / 测绘 / 物流 / 普查从业者的效率神器**
免费开源，本地运行，不限行数，双击安装就能用

</div>

---

## 😩 你是不是也天天被这些事折磨？

- 领导甩来3000条站点地址，要求全部转成经纬度坐标
- 在线工具免费版一次只能传100条，传5000条要点50次
- 高德Key一天5000次额度，跑到一半就超限报错
- 地址转出来飘到隔壁市，到了现场才发现错了
- 同一个地址反复查，API额度不知不觉就没了
- 想抓某个区域所有POI，还得自己写爬虫？

**这些坑我全踩过。然后我写了这个工具。**

---

## ✨ 三大核心功能

### 🗺️ 批量地址 → 经纬度（正向地理编码）

Excel丢一堆地址进去，一键全部跑出坐标，自动写回原表。

- ✅ 几千到几万行批量跑，不卡死
- ✅ 高德+百度双引擎，**高德额度用完自动切百度**，不用手动换Key
- ✅ 本地SQLite缓存，**查过的地址永久存，第二次查秒出不耗额度**
- ✅ 区县漂移**自动标黄**，问题行一眼锁定
- ✅ 跑完自动拆7列结构化地址：省 / 市 / 区县 / 街道 / 道路 / 门牌 / 小区

### 📍 批量经纬度 → 标准地址（逆向地理编码）

反过来用！手里有GPS坐标？批量反查完整标准地址。

- ✅ 同样批量处理，同样写回Excel
- ✅ 同样双引擎自动切换 + 本地缓存
- ✅ 输出完整结构化地址，做台账直接用

### 📡 POI 网格化地理信息采集

地图上框一个范围，工具自动按格子遍历，把区域内所有POI的**名称 + 坐标**全抓下来。

- ✅ 网格化自动遍历，不用手动一个个点
- ✅ 并发采集，速度快
- ✅ 自动去重（空间索引，同一个POI不重复抓）
- ✅ 关键词自动扩展（配了AI大模型后，"学校"自动展开大学/学院/中学/小学…）
- ✅ 结果直接导出Excel

**适用场景**：通信覆盖分析 · 选址调研 · 商圈分析 · 设施普查

---

## 💡 为什么选这个工具？

| 对比项 | 在线工具 | 手工查 | **本工具** |
| --- | --- | --- | --- |
| 行数限制 | 免费版限行数 | 无 | **无限制** |
| 费用 | 会员贵 | 免费 | **完全免费开源** |
| 双引擎切换 | 不支持 | 手动换 | **自动切换不中断** |
| 本地缓存 | 没有 | 没有 | **有，不重复耗额度** |
| 漂移检测 | 没有 | 肉眼看 | **自动标黄** |
| POI采集 | 另收费 | 写爬虫 | **内置，框选就抓** |
| 结构化拆列 | 手动拆 | 手动拆 | **自动7列** |
| Excel大文件 | 卡 | — | **秒开不卡** |

---

## 🚀 快速开始

### 方式一：下载安装包（推荐）

1. 去 [Releases](https://github.com/Mine1178/addressmap-tool/releases) 下载最新版 `setup_addressmap_tool.exe`
2. 双击安装，一键下一步
3. 启动后在「API密钥配置」页填入你自己的高德Key和百度AK
4. 导入Excel，开跑！

### 方式二：源码运行

```bash
git clone https://github.com/Mine1178/addressmap-tool.git
cd addressmap-tool
pip install openpyxl Pillow
python addressmap_tool_v2.8.66.py
```

### 🔑 密钥申请（免费）

| 平台 | 申请地址 | 需要开通 |
| --- | --- | --- |
| 高德开放平台 | https://lbs.amap.com/ | Web服务（地理/逆地理编码、POI搜索） |
| 百度地图开放平台 | https://lbsyun.baidu.com/ | 服务端AK |

> 💡 个人开发者每天都有几千次免费额度，日常办公完全够用。

---

## 📋 Excel 格式约定

使用 `Standard_Address_Template.xlsx` 模板，列说明：

| 列号 | 字段 | 必填 | 说明 |
| --- | --- | --- | --- |
| 1 | 序号 | 推荐 | 任意编号 |
| 2 | 省 | 可选 | 如「云南省」 |
| 3 | 市 | 可选 | 如「曲靖市」 |
| 4 | 区县 | 可选 | 参与查询拼接与漂移比对 |
| 5 | 地址1 | **必填** | 主地址 |
| 6 | 地址2 | 选填 | 地址1异常时的备用地址 |
| 7-8 | GCJ-02 经纬度 | 输出 | 正向编码结果 |
| 9-10 | WGS-84 经纬度 | 输出 | 坐标互转结果 |
| 11 | 标准地址 | 输出 | 标准化地址（漂移时标注） |

---

## 🛡️ 安全与合规

- API密钥仅保存在本地 `addr_config.json`，**不上传、不公开**
- 请遵守高德/百度开放平台服务条款与配额限制
- 本项目按 **Apache-2.0** 许可证开源，随便用、随便改
- 数据结果仅供参考，不构成任何测绘、法律或商业决策依据

---

## 🤝 贡献与反馈

- 发现Bug？提 [Issue](https://github.com/Mine1178/addressmap-tool/issues)
- 想加功能？欢迎 PR
- 觉得好用？**点个Star⭐支持一下**，作者还在持续更新！

---

<div align="center">

**如果这个工具帮你少加了班，欢迎请作者喝杯咖啡 ☕**

<br>

<img src="assets/alipay.jpg" width="300" alt="支付宝赞助">
&nbsp;&nbsp;&nbsp;&nbsp;
<img src="assets/wechatpay.jpg" width="300" alt="微信赞助">

<br>

Made with ❤️ by 通信勘察打工人

</div>
