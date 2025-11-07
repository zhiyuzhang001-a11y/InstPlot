# license_manager_secure.py

import hashlib, os, json, datetime, uuid, platform

LICENSE_FILE = "license.dat"
TRIAL_DAYS = 0
# 复杂的密钥：包含大小写字母、数字、特殊字符，长度至少32位
SECRET_KEY = "Xk9#mP2$vL8@wQ5!nR7&jT4*bN6^hG3%cF2"

def get_machine_code():
    """
    生成基于多个硬件信息的机器码
    包括: 计算机名、MAC地址、CPU信息
    """
    # 计算机名
    hostname = os.getenv("COMPUTERNAME") or os.getenv("HOSTNAME") or "Unknown"
    
    # MAC地址（更难伪造）
    try:
        mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) 
                       for elements in range(0,2*6,2)][::-1])
    except:
        mac = "Unknown"
    
    # CPU信息
    try:
        cpu = platform.processor() or "Unknown"
    except:
        cpu = "Unknown"
    
    # 组合所有信息生成唯一机器码
    combined = f"{hostname}_{mac}_{cpu}"
    return hashlib.sha256(combined.encode()).hexdigest()[:12].upper()

def get_license_key(machine_code):
    return hashlib.sha256((machine_code + SECRET_KEY).encode()).hexdigest()[:12].upper()

def read_license():
    """读取并验证许可证文件的完整性"""
    if not os.path.exists(LICENSE_FILE):
        return {}
    try:
        with open(LICENSE_FILE, "r") as f:
            data = json.load(f)
        
        # 验证文件完整性（防止手动修改）
        if "checksum" in data:
            content = {k: v for k, v in data.items() if k != "checksum"}
            expected_checksum = hashlib.sha256(
                (json.dumps(content, sort_keys=True) + SECRET_KEY).encode()
            ).hexdigest()
            if data["checksum"] != expected_checksum:
                # 文件被篡改，清空数据
                return {}
        return data
    except:
        return {}

def write_license(data):
    """写入许可证文件并添加完整性校验"""
    # 添加校验和防止手动修改
    content = {k: v for k, v in data.items() if k != "checksum"}
    checksum = hashlib.sha256(
        (json.dumps(content, sort_keys=True) + SECRET_KEY).encode()
    ).hexdigest()
    data["checksum"] = checksum
    
    with open(LICENSE_FILE, "w") as f:
        json.dump(data, f)

def check_license():
    """
    检查许可证状态
    返回: (is_valid, remaining_days)
    - is_valid: True=可以使用, False=不能使用
    - remaining_days: 剩余试用天数(0表示已激活或已过期)
    """
    data = read_license()
    
    # 检查是否已激活
    if "license_key" in data:
        # 已激活，验证授权码
        if data["license_key"] == get_license_key(get_machine_code()):
            return True, 0
        else:
            # 授权码无效（可能是篡改或机器码变化）
            return False, 0
    
    # 试用模式
    elif "trial_start" in data:
        try:
            start = datetime.datetime.fromisoformat(data["trial_start"])
            now = datetime.datetime.now()
            
            # 防止时间回调攻击
            if "last_check" in data:
                last = datetime.datetime.fromisoformat(data["last_check"])
                if now < last:
                    # 系统时间被往回调整，视为异常
                    return False, -1
            
            # 更新最后检查时间
            data["last_check"] = now.isoformat()
            write_license(data)
            
            # 计算剩余天数
            days_used = (now - start).days
            remaining = TRIAL_DAYS - days_used
            
            if remaining > 0:
                return True, remaining
            else:
                return False, 0
        except:
            # 数据格式错误，重置试用
            write_license({"trial_start": datetime.datetime.now().isoformat()})
            return True, TRIAL_DAYS
    
    else:
        # 首次使用，进入试用模式
        write_license({
            "trial_start": datetime.datetime.now().isoformat(),
            "last_check": datetime.datetime.now().isoformat()
        })
        return True, TRIAL_DAYS

def activate_app():
    """
    激活应用程序
    返回: True=激活成功, False=激活失败或取消
    """
    from PySide6.QtWidgets import QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox
    
    machine_code = get_machine_code()
    
    # 创建自定义对话框
    dialog = QDialog(None)
    dialog.setWindowTitle("软件激活")
    dialog.setMinimumWidth(450)
    
    layout = QVBoxLayout()
    
    # 标题
    title_label = QLabel("激活 InstPlot")
    title_label.setStyleSheet("font-size: 14pt; font-weight: bold; margin-bottom: 10px;")
    layout.addWidget(title_label)
    
    # 机器码标签
    label1 = QLabel("您的机器码（可直接复制）：")
    label1.setStyleSheet("font-size: 11pt; margin-top: 5px;")
    layout.addWidget(label1)
    
    # 机器码输入框（只读但可复制）
    machine_code_edit = QLineEdit(machine_code)
    machine_code_edit.setReadOnly(True)
    machine_code_edit.setStyleSheet(
        "font-size: 13pt; font-weight: bold; padding: 8px; "
        "background-color: #E8F5E9; border: 2px solid #4CAF50; "
        "border-radius: 5px; color: #1B5E20;"
    )
    machine_code_edit.selectAll()  # 默认全选
    layout.addWidget(machine_code_edit)
    
    # 说明
    info_label = QLabel("请将机器码发送给作者以获取授权码")
    info_label.setStyleSheet("font-size: 10pt; color: #666; margin-bottom: 15px;")
    layout.addWidget(info_label)
    
    # 授权码标签
    label2 = QLabel("请输入授权码：")
    label2.setStyleSheet("font-size: 11pt; margin-top: 10px;")
    layout.addWidget(label2)
    
    # 授权码输入框
    license_key_edit = QLineEdit()
    license_key_edit.setPlaceholderText("请输入12位授权码")
    license_key_edit.setStyleSheet(
        "font-size: 12pt; padding: 8px; border: 2px solid #2196F3; "
        "border-radius: 5px;"
    )
    layout.addWidget(license_key_edit)
    
    # 按钮
    button_layout = QHBoxLayout()
    button_layout.addStretch()
    
    btn_ok = QPushButton("确定")
    btn_ok.setStyleSheet(
        "font-size: 11pt; padding: 8px 25px; background-color: #4CAF50; "
        "color: white; border-radius: 5px; font-weight: bold;"
    )
    
    btn_cancel = QPushButton("取消")
    btn_cancel.setStyleSheet(
        "font-size: 11pt; padding: 8px 25px; background-color: #9E9E9E; "
        "color: white; border-radius: 5px;"
    )
    
    button_layout.addWidget(btn_ok)
    button_layout.addWidget(btn_cancel)
    layout.addLayout(button_layout)
    
    dialog.setLayout(layout)
    
    # 按钮事件
    btn_ok.clicked.connect(dialog.accept)
    btn_cancel.clicked.connect(dialog.reject)
    
    ret = dialog.exec()
    
    if ret == QDialog.Accepted:
        key = license_key_edit.text().strip()
        if key == get_license_key(machine_code):
            write_license({"license_key": key})
            QMessageBox.information(None, "激活成功", "软件已成功激活！")
            return True
        else:
            QMessageBox.warning(None, "错误", "授权码无效！\n请检查是否输入正确。")
            return False
    return False