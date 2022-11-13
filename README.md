**Windows 文件权限ACL列表查询导出工具, 详见工具介绍及参数示例.**

    Windows 文件权限ACL列表查询导出工具, 详见工具介绍及参数示例.

    --- 仅支持导出 Excel|Json 格式
    --- 支持按层级递归遍历文件夹, 将使用CLI携带的凭证
    --- 支持对文件夹进行递归的SID精确搜索和正则搜索  --- 正则通配符部分描述:  * 匹配0到多个   + 匹配1个或多个   ? 匹配0个或1个



    --------------------------------------- Author  : ww1372247148@163.com
    --------------------------------------- Version : 0.0.3-beta


```bash
optional arguments:
  -h, --help            show this help message and exit

Normal Options:
  PathName              路径名称 --- 输入需要递归遍历的文件夹路径, 支持绝对路径和相对路径

  -d DEPTH, --depth DEPTH
                        递归层级 - 默认:0 - 输入递归遍历的层级, 仅支持小于6的层级递归.  0 仅查询PathName的权限ACL   n (n<6) 递归遍历到n层级


Export Options:
  -w WRITE, --write WRITE
                        写入路径  --- 将查询结果写入到指定路径

  -t {xlsx,json}, --type {xlsx,json}
                        ACL格式  --- 将查询结果转换成 Excel|Json 格式


Beta Options:
  -fs FINDSID, --findsid FINDSID
                        用户搜索 --- 对用户SID进行查询, 使用精确搜索

  -reg-fs REG__FINDSID, --reg--findsid REG__FINDSID
                        用户搜索 --- 对用户SID进行查询, 使用正则搜索

  -p {default,tree,uacl,aacl,path} [{default,tree,uacl,aacl,path} ...], --print {default,tree,uacl,aacl,path} [{default,tree,uacl,aacl,path} ...]
                        过程回显 --- 按特定格式进行过程回显, 默认不开启过程回显, 只回显执行结果   (该选项必须放在最后输入)



    --- 递归遍历文件夹权限ACL, 按层级递归, 默认层级为 0 (层级 0 即仅查询PathName, 非递归)
    eg.=    <EXE> C:\ConfigMsi                                             # 仅查询 C:\ConfigMsi 的权限ACL.
    eg.=    <EXE> "C:\Config Msi" --depth=0                                # 长选项: 仅查询 C:\ConfigMsi 的权限ACL, 回显执行结果
    eg.=    <EXE> "C:\Config Msi" -d 2                                     # 短选项: 递归遍历 C:\ConfigMsi 子文件夹及文件的权限ACL, 回显执行结果, 层级为 2


    --- 查询 C:\ConfigMsi 的权限ACL, 导出为 Excel|Json 格式
    eg.=    <EXE> C:\ConfigMsi --write="C:\excel.xlsx" --type=xlsx         # 长选项: Excel格式, 绝对路径. 文件后缀必须匹配xlsx
    eg.=    <EXE> "C:\Config Msi" -w json.json -t json                     # 短选项: 文件夹路径有空格必须使用""包含. json格式, 相对路径.  文件后缀必须匹配json


    --# Beta功能: 查询文件夹权限ACL, 对用户名称进行搜索, 精确搜索和正则搜索的选项只能二选一
    --# 正则表达式:   * 匹配0到多个   + 匹配1个或多个   ? 匹配0个或1个
    eg.=    <EXE> C:\ConfigMsi --findsid=100001                            # 长选项: 精准搜索
    eg.=    <EXE> "C:\Config Msi" -d 2 -fs "*01*"                          # 短选项: 精准搜索
    eg.=    <EXE> "C:\Config Msi" -d 2 --reg--findsid "*01+"               # 长选项: 正则搜索
    eg.=    <EXE> "C:\Config Msi" -d 2 -reg-fs "*0+1?"                     # 短选项: 正则搜索


    --# Beta功能: 特定格式进行过程回显, 默认选项和其他选项只能二选一 ( -p 选项必须放在位置参数的最后 )
    --# 默认选项  回显执行结果: default
    --# 其他选项  树深度: tree   ACL用户: uacl   ACL权限: aacl   当前路径: path
    eg.=    <EXE> "C:\Config Msi" --print default                          # 长选项: 过程回显, 格式: default
    eg.=    <EXE> "C:\Config Msi" -d 2 -reg-fs "011*" -p uacl aacl path    # 短选项: 正则搜索, 过程回显, 格式: uacl aacl path
    eg.=    <EXE> "C:\Config Msi" -d 2 -p tree uacl aacl path              # 短选项: 过程回显, 格式: tree uacl aacl path
```
