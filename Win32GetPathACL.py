#!/bash/bin

import argparse
import ctypes
import os
import re

import utils
# ctypes.windll.shell32.IsUserAnAdmin()

# 变量声明
flag_print: int
'''
    0: 输出执行结果
    1: 输出default格式
    2: 输出other*格式 (从*args获取)
'''
flag_findsid: int
'''
    0: 不搜索
    1: 精准搜索
    2: 正则搜索
'''
flag_export: int
'''
    0: 不导出
    1: 导出Excel
    2: 导出Json
'''

# auths_arr格式声明
auths_arr = [{
    'id': int,  # 权限条目的序号
    'alias': str,  # 文件夹的别名
    'subDirs': str,  # 遍历当前文件夹的子文件夹.     -1 非文件夹 | -2 拒绝访问
    'subFiles': str,  # 遍历当前文件夹的文件.       -1 非文件夹 | -2 拒绝访问
    'domain': str,  # 所属组/域. 非组/域则为空
    'user': str,  # 用户名称
    'fullAccessMask': str,  # 完整的权限掩码, 包含继承关系. 参数详见icacls官方文档
    'accessMask': str,  # 不包含继承关系的权限掩码. 参数详见icacls官方文档
    'inherit': str,  # 继承关系. 参数详见icacls官方文档
    'isAllow': int,  # 允许或拒绝的权限. 参数: 1允许|0拒绝
    'parentInherit': int,  # 从父对象继承. 参数: 1继承|0不继承
    'propagateInherit': int,  # 递归传播权限. 参数: 1开启|0关闭
    'inheritRight': int,  # 权限的应用场景. 参数详见inheritRight_map
    'fullControl': int,  # 完全控制权限, 参数: 1开启|0关闭 
    'modify': int,  # 修改权限, 参数: 1开启|0关闭  (普通权限声明)
    'read_execute': int,  # 读取和执行, 参数: 1开启|0关闭  (普通权限声明)
    'read_only': int,  # 读取权限, 参数: 1开启|0关闭  (普通权限声明)
    'write_only': int,  # 写入权限, 参数: 1开启|0关闭  (普通权限声明)
    'readData_listDir': int,  # 列出文件夹/读取数据, 参数: 1开启|0关闭  (特殊权限声明)
    'readAttr': int,  # 读取属性, 参数: 1开启|0关闭  (特殊权限声明)
    'readExtAttr': int,  # 读取扩展属性, 参数: 1开启|0关闭  (特殊权限声明)
    'readPermiss': int,  # 读取权限, 参数: 1开启|0关闭  (特殊权限声明)
    'execute_traverse': int,  # 遍历文件夹/执行文件, 参数: 1开启|0关闭  (特殊权限声明)
    'writeData_addFile': int,  # 创建文件/写入数据, 参数: 1开启|0关闭  (特殊权限声明)
    'appendData_addSubdir': int,  # 创建文件夹/附加数据, 参数: 1开启|0关闭  (特殊权限声明)
    'writeAttr': int,  # 写入属性, 参数: 1开启|0关闭  (特殊权限声明)
    'writeExtAttr': int,  # 写入扩展属性, 参数: 1开启|0关闭  (特殊权限声明)
    'delete': int,  # 删除, 参数: 1开启|0关闭  (特殊权限声明)
    'deleteChild': int,  # 删除子文件夹及文件, 参数: 1开启|0关闭  (特殊权限声明)
    'changePermiss': int,  # 更改权限, 参数: 1开启|0关闭  (特殊权限声明)
    'takeOwner': int,  # 取得所有权, 参数: 1开启|0关闭  (特殊权限声明)
    'sync': int  # 同步, 参数: 1开启|0关闭  (特殊权限声明)
}]

# path_auths格式声明
path_auths = {
    'dirs':  # 文件夹的权限ACL列表
    {
        str: auths_arr  # 文件夹的绝对路径和权限ACL
    },
    'files':  # 文件的权限ACL列表
    {
        str: auths_arr  # 文件的绝对路径和权限ACL
    }
}

# inheritRight_map内容声明
inheritRight_map = {
    0: '',  # 只有该文件夹
    1: '(CI)(IO)',  # 只有子文件夹
    2: '(OI)(IO)',  # 只有文件
    3: '(CI)',  # 此文件夹和子文件夹
    4: '(OI)',  # 此文件夹和文件
    5: '(OI)(CI)(IO)',  # 仅子文件夹和文件
    6: '(OI)(CI)'  # 此文件夹、子文件夹和文件
}
'''
    0: 只有该文件夹
    1: 只有子文件夹
    2: 只有文件
    3: 此文件夹和子文件夹
    4: 此文件夹和文件
    5: 仅子文件夹和文件
    6: 此文件夹、子文件夹和文件
'''


def get_path_authority(path_: str, flag_findsid_: int, flag_print_: int, findsid_: str = '', print_format_: list = None):
    '''
        遍历path_的权限ACL
        path_           : 文件夹的绝对路径
        flag_findsid_   : 0 不搜索 | 1 精准搜索 | 2 正则搜索
        flag_print_     : 0 输出执行结果 | 1 输出default格式 | 2 输出other*格式 (从*args获取)
        findsid_        : 搜索该用户名称
        print_format_   : 格式输出的选项集.  default | [ tree, uacl, aacl, path ] 
    
    SUCCESS
    return dict         : 输出权限ACL的dict对象
    
    ERROR
    return dict(int:list): 输出报错的code和msg
    '''
    if not os.path.exists(path_):
        return {4004: ['query path not exists, please try again', '查询路径不存在, 请重新输入']}
    computerName = os.environ.get('computername')

    auth_id = 0
    auths_list = []
    try:
        if path_[-2:] == '\\\\':
            tmp_path_ = path_[:-1]
            shell_code = 'icacls "%s"' % path_
        else:
            tmp_path_ = path_
            if path_[-2] != '\\' and path_[-1] == '\\':
                shell_code = 'icacls "%s"' % path_.replace('\\', '\\\\')
            else:
                shell_code = 'icacls "%s"' % path_
        auth_subDirs = [os.path.join(tmp_path_, v) for v in os.listdir(tmp_path_)
                        if os.path.isdir(os.path.join(tmp_path_, v))] if os.path.isdir(tmp_path_) else -1
        auth_subFiles = [os.path.join(tmp_path_, v) for v in os.listdir(tmp_path_)
                         if os.path.isfile(os.path.join(tmp_path_, v))] if os.path.isdir(tmp_path_) else -1
    except WindowsError:
        auth_subDirs = -2
        auth_subFiles = -2
    with os.popen(shell_code, 'r') as auths:
        for line in auths.readlines():
            path_ = path_[:-1] if path_[-2:] == '\\\\' else path_
            if path_ in line:
                line = line.replace('\n', '').replace(path_, '').strip()
            else:
                line = line.replace('\n', '').strip()
                
            if '未设置任何权限。所有用户都具有完全控制权限。' in line:
                line = 'Everyone:(F)'
                
            if '处理 1 个文件时失败' in line:
                auths_list = '拒绝访问'
            elif line != '' and '已成功处理' not in line:
                auth_alias = os.path.basename(path_)
                auth_domain = line[line[:line.rfind('\\')].rfind(' ') + 1:line.rfind('\\')] if line.rfind('\\') > 0 else ''
                auth_user = line[line.rfind('\\') + 1:line.rfind(':')]

                try:
                    if (flag_findsid_ == 1 and findsid_ != auth_user) or (flag_findsid_ == 2 and re.match(r'%s' % findsid_, auth_user, flags=re.I) == None):
                        continue
                except Exception as e:
                    return {4000: [e, '正则表达式有误, 请重新输入']}

                auth_fullAccessMask = line[line.rfind(':') + 1:]
                auth_isAllow = 0 if '(DENY)' in auth_fullAccessMask or '(N)' in auth_fullAccessMask else 1
                auth_fullControl = 1 if '(F)' in auth_fullAccessMask or '(N)' in auth_fullAccessMask else 0

                if not auth_isAllow and auth_fullControl:
                    auth_inherit = ""
                else:
                    auth_inherit = str.join('', [
                        '%s)' % v for v in ((auth_fullAccessMask.split(')')[:-3] if '(DENY)' in auth_fullAccessMask else auth_fullAccessMask.split(')')[:-2]
                                             ) if len(auth_fullAccessMask.split(')')) > 2 else auth_fullAccessMask.split(')')[:-2])
                    ])
                auth_parentInherit = 1 if '(I)' in auth_inherit else 0
                auth_propagateInherit = 0 if '(NP)' in auth_inherit else 1
                auth_inheritRight = [k for k, v in inheritRight_map.items() if v == auth_inherit.replace('(I)', '').replace('(NP)', '')][0]
                auth_accessMask = auth_fullAccessMask.split(")")[-2][1:].upper().split(',')

                auth_simple_modify = 1 if len([v for v in auth_accessMask if 'M' == v]) == 1 else 0
                auth_simple_read_execute = 1 if auth_simple_modify or len([v for v in auth_accessMask if 'RX' == v]) == 1 else 0
                auth_simple_read_only = 1 if auth_simple_read_execute or len([v for v in auth_accessMask if 'R' == v]) == 1 else 0
                auth_simple_write_only = 1 if auth_simple_modify or len([v for v in auth_accessMask if 'W' == v]) == 1 else 0

                auth_special_readData_listDir = 1 if auth_simple_read_only or len([v for v in auth_accessMask if 'RD' == v]) == 1 else 0
                auth_special_readAttr = 1 if auth_simple_read_only or len([v for v in auth_accessMask if 'RA' == v]) == 1 else 0
                auth_special_readExtAttr = 1 if auth_simple_read_only or len([v for v in auth_accessMask if 'REA' == v]) == 1 else 0
                auth_special_readPermiss = 1 if auth_simple_read_only or len([v for v in auth_accessMask if 'RC' == v]) == 1 else 0
                auth_special_execute_traverse = 1 if auth_simple_read_execute or len([v for v in auth_accessMask if 'X' == v]) == 1 else 0
                auth_special_writeData_addFile = 1 if auth_simple_write_only or len([v for v in auth_accessMask if 'WD' == v]) == 1 else 0
                auth_special_appendData_addSubdir = 1 if auth_simple_write_only or len([v for v in auth_accessMask if 'AD' == v]) == 1 else 0
                auth_special_writeAttr = 1 if auth_simple_write_only or len([v for v in auth_accessMask if 'WA' == v]) == 1 else 0
                auth_special_writeExtAttr = 1 if auth_simple_write_only or len([v for v in auth_accessMask if 'WEA' == v]) == 1 else 0
                auth_special_delete = 1 if auth_simple_modify or len([v for v in auth_accessMask if 'D' == v]) == 1 else 0
                auth_special_deleteChild = 1 if len([v for v in auth_accessMask if 'DC' == v]) == 1 else 0
                auth_special_changePermiss = 1 if len([v for v in auth_accessMask if 'WDAC' == v]) == 1 else 0
                auth_special_takeOwner = 1 if len([v for v in auth_accessMask if 'WO' == v]) == 1 else 0
                auth_special_sync = 1 if len([v for v in auth_accessMask if 'S' == v]) == 1 else 0
                auths_list.append({
                    'id': auth_id,
                    'alias': auth_alias,
                    'domain': computerName if auth_domain == 'BUILTIN' else auth_domain,
                    'user': auth_user,
                    'fullAccessMask': auth_fullAccessMask,
                    'accessMask': auth_accessMask,
                    'inherit': auth_inherit,
                    'isAllow': auth_isAllow,
                    'parentInherit': auth_parentInherit,
                    'propagateInherit': auth_propagateInherit,
                    'inheritRight': auth_inheritRight,
                    'fullControl': auth_fullControl,
                    'modify': auth_simple_modify,
                    'read_execute': auth_simple_read_execute,
                    'read_only': auth_simple_read_only,
                    'write_only': auth_simple_write_only,
                    'readData_listDir': auth_special_readData_listDir,
                    'readAttr': auth_special_readAttr,
                    'readExtAttr': auth_special_readExtAttr,
                    'readPermiss': auth_special_readPermiss,
                    'execute_traverse': auth_special_execute_traverse,
                    'writeData_addFile': auth_special_writeData_addFile,
                    'appendData_addSubdir': auth_special_appendData_addSubdir,
                    'writeAttr': auth_special_writeAttr,
                    'writeExtAttr': auth_special_writeExtAttr,
                    'delete': auth_special_delete,
                    'deleteChild': auth_special_deleteChild,
                    'changePermiss': auth_special_changePermiss,
                    'takeOwner': auth_special_takeOwner,
                    'sync': auth_special_sync
                })
                auth_id += 1
                print(
                    '树深度: {}'.format(path_.count('\\')).ljust(8, ' ') if 'tree' in print_format_ or 'default' in print_format_ else '',
                    'ACL用户:' if 'uacl' in print_format_ or 'default' in print_format_ else '',
                    '{_dm}{_split}{_uname}'.format(_dm=auth_domain, _split='\\' if auth_domain else '', _uname=auth_user).ljust(
                        38 -
                        len(str.join('', re.findall(r'[\u4e00-\u9fa5]+', auth_user))), ' ') if 'uacl' in print_format_ or 'default' in print_format_ else '',
                    'ACL权限: {}'.format(auth_fullAccessMask) if 'aacl' in print_format_ or 'default' in print_format_ else '',
                    '{}:'.format('文件夹' if os.path.isdir(path_) else '文件  ').rjust(
                        ((38 if os.path.isdir(path_) else 39) -
                         len(auth_fullAccessMask)) if 'path' not in print_format_ or 'aacl' in print_format_ or 'default' in print_format_ else 0, ' ')
                    if 'path' in print_format_ or 'default' in print_format_ else '',
                    path_ if 'path' in print_format_ or 'default' in print_format_ else '') if flag_print_ > 0 else ''

                # print(
                #     '树深度: %s'.ljust(8, ' ') % path_.count('\\'), '{_ph}: {_pname}'.format(_ph='文件夹' if os.path.isdir(path_) else '文件  ', _pname=path_),
                #     'ACL用户:'.rjust(
                #         130 - len(path_) -
                #         (0 if len(re.findall(r'[\u4e00-\u9fa5]+', path_)) == 0 else len(str.join('', re.findall(r'[\u4e00-\u9fa5]+', path_)))), ' '),
                #     '{_dm}{_split}{_uname}'.format(_dm=auth_domain, _split='\\' if auth_domain else '', _uname=auth_user), 'ACL权限: %s'.rjust(
                #         55 - len(auth_domain + '\\' + auth_user if auth_domain else auth_user) - len(str.join('', re.findall(r'[\u4e00-\u9fa5]+', auth_user))),
                #         ' ') % auth_fullAccessMask)
        auth_usersList = list(
            set([('{}\\{}'.format(item['domain'], item['user']) if item['domain'] != '' else item['user']) for item in auths_list if auths_list != '拒绝访问']))
    return {path_: {'accessState': auths_list, 'subDirs': auth_subDirs, 'subFiles': auth_subFiles, 'count_result': [auth_usersList]}}


def loop_get_walks(paths_: str,
                   flag_findsid_: int,
                   flag_print_: int,
                   depthLevel_: int = 2,
                   write_path_: str = '',
                   write_type_: str = '',
                   findsid_: str = '',
                   print_format_: list = None):
    '''
        递归遍历paths_的权限ACL
        paths_          : 文件夹的绝对路径
        flag_findsid_   : 0 不搜索 | 1 精准搜索  | 2 正则搜索
        flag_print_     : 0 输出执行结果 | 1 输出default格式 | 2 输出other*格式
        depthLevel_     : 递归的树深度.  0遍历当前目录 | 1 递归遍历到第一级目录 | 2 递归遍历到第二级目录 | 3...以此类推
        write_path_     : 写入路径
        wrte_type_      : 写入格式.  Excel表格: xlsx  |  Json文件: json
        findsid_        : 搜索该用户名称
        print_format_   : 格式输出的选项集.  default | [ tree, uacl, aacl, path ] 
    
    SUCCESS
    return dict,dict    : 输出权限ACL的dict对象
    
    ERROR
    return int,list      : 报错返回code和错误信息
    '''
    if not os.path.exists(paths_):
        return 4004, ['query path not exists, please try again', '查询路径不存在, 请重新输入']
    elif int != type(depthLevel_) or not str(depthLevel_).isdecimal() or not 0 <= depthLevel_ < 6:
        return 4011, ['depth format error, please enter number 0-5', 'depth格式错误, 请输入数字 0-5']

    write_path_ = '' if not write_path_ else write_path_
    write_type_ = '' if not write_type_ else write_type_
    findsid_ = '' if not findsid_ else findsid_
    map_printFormat = {'tree': '树深度', 'uacl': 'ACL用户', 'aacl': 'ACL权限', 'path': '当前路径'}
    trans_printFormat = [v for k, v in map_printFormat.items() if k in str.join(' ', print_format_)] if print_format_ and print_format_[0] != 'default' else ''
    parse_print = '仅结果' if flag_print_ == 0 else ('default格式' if flag_print_ == 1 else str.join(' ', trans_printFormat))
    print('当前配置 ⬇⬇⬇  (注意: 递归层级越大, 等待时间越长.)')
    print('查询路径: {}'.format(paths_))
    print('递归层级: {_depth}{_n}格式输出: {_format}{_n}{_reg}搜索用户: {_findsid}'.format(_path=paths_,
                                                                               _depth=depthLevel_,
                                                                               _format=parse_print,
                                                                               _findsid=findsid_,
                                                                               _n=' ' * 10,
                                                                               _reg='(正则表达式)' if flag_findsid_ == 2 else ''))
    print('写入路径: {}'.format(write_path_))
    print('写入格式: {}\n\n'.format(write_type_))

    if depthLevel_ == 0:
        pathAuths = get_path_authority('%s\\' % paths_ if paths_[-2:] == ':\\' else paths_, flag_findsid_, flag_print_, findsid_, print_format_)
        for k, v in pathAuths.items():
            if type(k) == int:
                return k, v
        count_success_dirs = 0
        count_success_files = 0
        count_fail_dirs = 0
        count_fail_files = 0
        auth_fail_path_list = []
        for k, v in pathAuths.items():
            list_authUser = v['count_result'][0]
            if v['subDirs'] != -1:
                count_success_dirs += 1 if v['accessState'] != '拒绝访问' else 0
                count_fail_dirs += 1 if v['accessState'] == '拒绝访问' else 0
                auth_fail_path_list.append(k) if v['accessState'] == '拒绝访问' else ''
            elif v['subDirs'] == -1:
                count_success_files += 1 if v['accessState'] != '拒绝访问' else 0
                count_fail_files += 1 if v['accessState'] == '拒绝访问' else 0
                auth_fail_path_list.append(k) if v['accessState'] == '拒绝访问' else ''
            del v['count_result']
        list_authUser = list(set(list_authUser))
        count_authUser = len(list_authUser)
        auth_fail_reason_list = ['拒绝访问'] if len(auth_fail_path_list) > 0 else []

        # 输出执行结果
        print('\n\n查询完毕. 结果反馈 ⬇⬇⬇')
        print('1. 查询到的用户数量: %d' % count_authUser)
        print('2. 查询到的用户名称: \n%s' % str.join(', ', list_authUser))
        print('3. 成功处理文件夹数量: %s' % count_success_dirs)
        print('4. 成功处理文件数量: %s' % count_success_files)
        print('5. 处理失败的 文件/文件夹 数量: %s' % (count_fail_dirs + count_fail_files))
        print('6. 处理失败的原因集合: %s' % str.join(', ', auth_fail_reason_list)) if (count_fail_dirs + count_fail_files) != 0 else ''
        print('7. 处理失败的 文件/文件夹 列表: \n%s' % str.join(', ', auth_fail_path_list)) if (count_fail_dirs + count_fail_files) != 0 else ''
        count_result = {
            'count_authUser': count_authUser,
            'list_authUser': list_authUser,
            'count_success_dirs': count_success_dirs,
            'count_success_files': count_success_files,
            'count_fail_dirs': count_fail_dirs,
            'count_fail_files': count_fail_files,
            'auth_fail_reason_list': auth_fail_reason_list,
            'auth_fail_path_list': auth_fail_path_list
        }
        return pathAuths, count_result
    else:
        walks = utils.depth_walk(paths_, depthLevel_)
        pathAuths = {'dirs': {}, 'files': {}}
        list_authUser = []
        for lWalk in walks:
            for dir in lWalk['dirs']:
                if '$RECYCLE.BIN' not in dir:
                    dir = os.path.join(lWalk['root'], dir)
                    for k, v in get_path_authority(dir, flag_findsid_, flag_print_, findsid_, print_format_).items():
                        if type(k) == int:
                            return k, v
                        pathAuths['dirs'][k] = {'accessState': v['accessState'], 'subDirs': v['subDirs'], 'subFiles': v['subFiles']}
                        [list_authUser.append(item) for item in v['count_result'][0]]
                else:
                    continue
            for file in lWalk['files']:
                file = os.path.join(lWalk['root'], file)
                for k, v in get_path_authority(file, flag_findsid_, flag_print_, findsid_, print_format_).items():
                    if type(k) == int:
                        return k, v
                    pathAuths['files'][k] = {'accessState': v['accessState'], 'subDirs': v['subDirs'], 'subFiles': v['subFiles']}
                    [list_authUser.append(item) for item in v['count_result'][0]]
        list_authUser = list(set(list_authUser))
        count_authUser = len(list_authUser)

        count_success_dirs = 0
        count_success_files = 0
        count_fail_dirs = 0
        count_fail_files = 0
        auth_fail_path_list = []
        for k1, v1 in pathAuths.items():
            for k2, v2 in v1.items():
                if k1 == 'dirs':
                    count_success_dirs += 1 if v2['accessState'] != '拒绝访问' else 0
                    count_fail_dirs += 1 if v2['accessState'] == '拒绝访问' else 0
                    auth_fail_path_list.append(k2) if v2['accessState'] == '拒绝访问' else ''
                elif k1 == 'files':
                    count_success_files += 1 if v2['accessState'] != '拒绝访问' else 0
                    count_fail_files += 1 if v2['accessState'] == '拒绝访问' else 0
                    auth_fail_path_list.append(k2) if v2['accessState'] == '拒绝访问' else ''
        auth_fail_reason_list = ['拒绝访问'] if len(auth_fail_path_list) > 0 else []
        # 输出执行结果
        print('\n\n查询完毕. 结果反馈 ⬇⬇⬇')
        print('1. 查询到的用户数量: %d' % count_authUser)
        print('2. 查询到的用户名称: \n%s' % str.join(', ', list_authUser))
        print('3. 成功处理文件夹数量: %s' % count_success_dirs)
        print('4. 成功处理文件数量: %s' % count_success_files)
        print('5. 处理失败的 文件/文件夹 数量: %s' % (count_fail_dirs + count_fail_files))
        print('6. 处理失败的原因集合: %s' % str.join(', ', auth_fail_reason_list))
        print('7. 处理失败的 文件/文件夹 列表: \n%s' % str.join(', ', auth_fail_path_list))
        count_result = {
            'count_authUser': count_authUser,
            'list_authUser': list_authUser,
            'count_success_dirs': count_success_dirs,
            'count_success_files': count_success_files,
            'count_fail_dirs': count_fail_dirs,
            'count_fail_files': count_fail_files,
            'auth_fail_reason_list': auth_fail_reason_list,
            'auth_fail_path_list': auth_fail_path_list
        }
        return pathAuths, count_result


def start_program():
    '''
        Windows 文件权限ACL列表查询导出工具, 详见工具介绍及参数示例.
    '''
    # 1. 创建解释器
    parser = argparse.ArgumentParser(usage='输入 "%(prog)s --help" 获取更多帮助信息',
                                     description='''
        - Copyright (c) 2022 uwellit.com, All Rights Reserved.
        -
        - Licensed under the PSF License;
        - you may not use this file except in compliance with the License. You may obtain a copy of the License at
        -
        - Https://docs.python.org/zh-cn/3/license.html#psf-license
        - 
        - Unless required by applicable law or agreed to in writing,
        - software distributed under the License is distributed on an "AS IS" BASIS,
        - WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
        - See the License for the specific language governing permissions and limitations under the License.
        

    Windows 文件权限ACL列表查询导出工具, 详见工具介绍及参数示例.
    --- 仅支持导出 Excel|Json 格式
    --- 支持按层级递归遍历文件夹, 将使用CLI携带的凭证
    --- 支持对文件夹进行递归的SID精确搜索和正则搜索  --- 正则通配符部分描述:  * 匹配0到多个   + 匹配1个或多个   ? 匹配0个或1个
    
    
    
    
    --------------------------------------- Author  : ww1372247148@163.com
    --------------------------------------- Version : 0.0.3-beta
    ''',
                                     epilog='''
    
    --- 递归遍历文件夹权限ACL, 按层级递归, 默认层级为 0 (层级 0 即仅查询PathName, 非递归)
    eg.=    <EXE> C:\\ConfigMsi                                             # 仅查询 C:\\ConfigMsi 的权限ACL.
    eg.=    <EXE> "C:\\Config Msi" --depth=0                                # 长选项: 仅查询 C:\\ConfigMsi 的权限ACL, 回显执行结果
    eg.=    <EXE> "C:\\Config Msi" -d 2                                     # 短选项: 递归遍历 C:\\ConfigMsi 子文件夹及文件的权限ACL, 回显执行结果, 层级为 2
    \n
    --- 查询 C:\\ConfigMsi 的权限ACL, 导出为 Excel|Json 格式
    eg.=    <EXE> C:\\ConfigMsi --write="C:\\excel.xlsx" --type=xlsx         # 长选项: Excel格式, 绝对路径. 文件后缀必须匹配xlsx
    eg.=    <EXE> "C:\\Config Msi" -w json.json -t json                     # 短选项: 文件夹路径有空格必须使用""包含. json格式, 相对路径.  文件后缀必须匹配json
    \n
    --# Beta功能: 查询文件夹权限ACL, 对用户名称进行搜索, 精确搜索和正则搜索的选项只能二选一
    --# 正则表达式:   * 匹配0到多个   + 匹配1个或多个   ? 匹配0个或1个
    eg.=    <EXE> C:\\ConfigMsi --findsid=100001                            # 长选项: 精准搜索 
    eg.=    <EXE> "C:\\Config Msi" -d 2 -fs "*01*"                          # 短选项: 精准搜索
    eg.=    <EXE> "C:\\Config Msi" -d 2 --reg--findsid "*01+"               # 长选项: 正则搜索
    eg.=    <EXE> "C:\\Config Msi" -d 2 -reg-fs "*0+1?"                     # 短选项: 正则搜索
    \n
    --# Beta功能: 特定格式进行过程回显, 默认选项和其他选项只能二选一 ( -p 选项必须放在位置参数的最后 )
    --# 默认选项  回显执行结果: default
    --# 其他选项  树深度: tree   ACL用户: uacl   ACL权限: aacl   当前路径: path
    eg.=    <EXE> "C:\\Config Msi" --print default                          # 长选项: 过程回显, 格式: default
    eg.=    <EXE> "C:\\Config Msi" -d 2 -reg-fs "011*" -p uacl aacl path    # 短选项: 正则搜索, 过程回显, 格式: uacl aacl path
    eg.=    <EXE> "C:\\Config Msi" -d 2 -p tree uacl aacl path              # 短选项: 过程回显, 格式: tree uacl aacl path

        ''',
                                     formatter_class=argparse.RawTextHelpFormatter)

    # 2. 创建分组
    normal_group = parser.add_argument_group(title='Normal Options')
    export_group = parser.add_argument_group(title='Export Options')
    beta_group = parser.add_argument_group(title='Beta Options')

    # 3. 添加参数到这些分组中
    normal_group.add_argument('PathName', type=str, help='''路径名称 --- 输入需要递归遍历的文件夹路径, 支持绝对路径和相对路径
    ''')
    normal_group.add_argument('-d',
                              '--depth',
                              type=int,
                              default=0,
                              help='''递归层级 - 默认:0 - 输入递归遍历的层级, 仅支持小于6的层级递归.  0 仅查询PathName的权限ACL   n (n<6) 递归遍历到n层级
    ''')
    # 参数解释
    # -d 代表短选项，在命令行输入-gf和--girlfriend的效果是一样的，作用是简化参数输入
    # --depth 代表完整的参数名称，可以尽量做到让人见名知意，需要注意的是如果想通过解析后的参数取出该值，必须使用带--的名称
    # type  代表输入的参数类型，从命令行输入的参数，默认是字符串类型
    # default 代表如果该参数不输入，则会默认使用该值
    # choices 代表输入参数的只能是这个choices里面的内容，其他内容则会保错

    beta_group.add_argument('-fs', '--findsid', type=str, help='''用户搜索 --- 对用户SID进行查询, 使用精确搜索
                            ''')
    beta_group.add_argument('-reg-fs', '--reg--findsid', type=str, help='''用户搜索 --- 对用户SID进行查询, 使用正则搜索
                            ''')
    beta_group.add_argument('-p',
                            '--print',
                            type=str,
                            nargs='+',
                            choices=['default', 'tree', 'uacl', 'aacl', 'path'],
                            help='''过程回显 --- 按特定格式进行过程回显, 默认不开启过程回显, 只回显执行结果   (该选项必须放在最后输入)
    ''')
    export_group.add_argument('-w', '--write', type=str, help='''写入路径  --- 将查询结果写入到指定路径
    ''')
    export_group.add_argument('-t', '--type', type=str, choices=['xlsx', 'json'], help='''ACL格式  --- 将查询结果转换成 Excel|Json 格式
    ''')

    # 4. 进行参数解析
    args = parser.parse_args()
    # args = parser.parse_args(['C:\\test', '-d=0', '-p=default'])

    # 5. 函数响应
    result = get_flag_argsParse(args)
    if result:
        flag_export, flag_findsid, flag_print = result
        if flag_export != 0 and args.write[args.write.rfind('.') + 1:] not in ['xlsx', 'json']:
            print('error: export file suffix mismatching, please try again')
            print('错 误: 文件后缀不匹配, 请重新输入')
            return
        # print(args)
        resultJson, _result = loop_get_walks(args.PathName, flag_findsid, flag_print, args.depth, args.write, args.type,
                                             args.findsid if flag_findsid == 1 else args.reg__findsid, args.print)
        if type(resultJson) == int:
            print('error: %s' % _result[0])
            print('错 误: %s' % _result[1])
        if flag_export != 0:
            if args.write.find(':\\') != -1:
                # 绝对路径
                utils.AuthsExport(resultJson, file_path_=args.write, file_export_=flag_export, kwargs=_result)
            else:
                # 相对路径
                utils.AuthsExport(resultJson, file_path_=os.path.join(os.getcwd(), args.write), file_export_=flag_export, kwargs={'count_result': _result})
    else:
        return


def get_flag_argsParse(args_: argparse.Namespace):
    '''
        判断用户输入的参数, 返回指定的选项标记
        args_       : 用户输入的参数集合
    
    SUCCESS
    return int,int,int  : 成功返回指定的选项标记
    
    ERROR
    return False        : 报错返回False
    '''
    # Normal + Export
    if args_.write != None and not args_.type:
        print('error: the following arguments are required: --type')
        print('错 误: 缺少 --type 参数, 请重新输入')
        return False
    elif args_.write == None and args_.type:
        print('error: the following arguments are required: --write')
        print('错 误: 缺少 --write 参数, 请重新输入')
        return False
    elif args_.write == None and not args_.type:
        flag_export = 0
    else:
        if args_.write.find(':\\') != -1:
            # 绝对路径
            if not os.path.isdir(os.path.dirname(args_.write)):
                print('error: directory not exists, please try again')
                print('错 误: 文件路径的目录不存在, 请重新输入')
                return False
            elif os.path.exists(args_.write):
                print('error: file already exists in "%s", please try again' % os.path.join(os.getcwd(), args_.write))
                print('错 误: "%s" 路径已存在文件, 请重新输入' % os.path.join(os.getcwd(), args_.write))
                return False
            else:
                flag_export = 1 if args_.type == 'xlsx' else 2
        else:
            # 相对路径
            if not os.path.isdir(os.path.dirname(os.path.join(os.getcwd(), args_.write))):
                print('error: directory not exists, please try again')
                print('错 误: 文件路径的目录不存在, 请重新输入')
                return False
            elif os.path.exists(os.path.join(os.getcwd(), args_.write)):
                print('error: file already exists in "%s", please try again' % os.path.join(os.getcwd(), args_.write))
                print('错 误: "%s" 路径已存在文件, 请重新输入' % os.path.join(os.getcwd(), args_.write))
                return False
            else:
                flag_export = 1 if args_.type == 'xlsx' else 2

    # Normal + Beta1
    if args_.findsid and args_.reg__findsid:
        print('error: only one option can be selected by between --findsid and --reg--findsid, please try again')
        print('错 误: --findsid和--reg--findsid的选项只能二选一, 请重新输入')
        return False
    elif args_.findsid == None and args_.reg__findsid == None:
        flag_findsid = 0
    elif args_.findsid:
        flag_findsid = 1
    else:
        flag_findsid = 2

    # Normal + Beta2
    if not args_.print:
        flag_print = 0
    else:
        if len(args_.print) == 1 and args_.print[0] == 'default':
            flag_print = 1
        elif 'default' not in args_.print:
            flag_print = 2
        else:
            print('error: default option should not in here, please try again')
            print('错 误: --print的default选项和其他选项只能二选一, 请重新输入')
            return False
    return (flag_export, flag_findsid, flag_print)


if __name__ == '__main__':
    try:
        start_program()
    except Exception as e:
        input('\n\n 程序出现BUG. 按任意键退出. <key>')
        os.close(0)
