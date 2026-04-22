import sys
import os
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTextEdit, QFileDialog, QMessageBox,
    QHBoxLayout, QVBoxLayout, QTableWidgetItem
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QFont
from qfluentwidgets import (
    PrimaryPushButton, PushButton, LineEdit, BodyLabel, TitleLabel,
    CheckBox, TableWidget, CardWidget, setTheme, Theme, FlowLayout
)

"""
抖音线索整理&合并工具

功能：
1. 批量处理和整理抖音线索Excel文件
2. 合并多个Excel文件并自动去重
3. 基于手机号和微信号的智能查重
4. 自定义本地线索判断逻辑
5. 现代简约风格界面
6. 自动时间戳命名
7. 处理完成后可选择直接打开生成的文件
"""

class ExcelTool(QMainWindow):
    """主应用类，继承自QMainWindow"""

    def __init__(self):
        """初始化应用程序"""
        super().__init__()
        # 设置窗口标题和大小
        self.setWindowTitle("抖音线索整理工具 v2.0 - By MirrorLab")
        self.resize(700, 1200)
        self.setMinimumSize(700, 1200)

        # 设置窗口图标
        logo_path = os.path.join(os.path.dirname(__file__), 'favicon.ico')
        if os.path.exists(logo_path):
            self.setWindowIcon(QIcon(logo_path))

        # 存储选中的文件
        self.selected_files = []
        
        # 存储处理过的源文件路径
        self.processed_files = []

        # 所有可用的列及其默认选中状态
        self.column_checkboxes = {
            "月份": True,          # 原"本月月份"
            "下发日期": True,
            "更新时间": True,
            "主播": True,
            "跟进顾问": True,
            "客户姓名/抖音ID": True,  # 默认选中
            "客户姓名": False,      # 默认不选
            "抖音ID": False,        # 默认不选
            "客户手机号": True,     # 原"手机号"
            "微信号": True,
            "客户所在地": False     # 默认不选
        }

        # 尝试使用桌面作为默认目录，如果失败则使用当前目录
        try:
            desktop_dir = os.path.join(os.path.expanduser('~'), 'Desktop')
            # 测试桌面目录是否有写入权限
            test_file = os.path.join(desktop_dir, ".test_permission.txt")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)
            self.output_dir = desktop_dir
        except PermissionError:
            # 如果桌面权限不足，使用当前目录
            self.output_dir = os.getcwd()

        # 设置应用主题为浅色模式
        setTheme(Theme.LIGHT)

        # 设置用户界面
        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        # 创建中央部件（无半透明效果）
        central_widget = QWidget()
        central_widget.setStyleSheet("QWidget { background-color: white; }")
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        title = TitleLabel("抖音线索整理工具")
        title.setFont(QFont("Caribi, Microsoft YaHei", 14, QFont.Bold))
        title.setStyleSheet("color: #000000;")
        main_layout.addWidget(title, alignment=Qt.AlignCenter)
        main_layout.addSpacing(15)

        file_card = CardWidget()
        file_card_layout = QVBoxLayout(file_card)
        file_card_layout.setContentsMargins(15, 15, 15, 15)
        file_card_layout.setSpacing(12)

        file_title = BodyLabel("文件选择")
        file_title.setFont(QFont("Caribi, Microsoft YaHei", 12, QFont.Bold))
        file_title.setStyleSheet("color: #000000;")
        file_card_layout.addWidget(file_title)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        select_file_btn = PrimaryPushButton("选择文件")
        select_file_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        select_file_btn.setFixedSize(190, 45)
        select_file_btn.clicked.connect(self.select_files)

        select_folder_btn = PushButton("选择文件夹")
        select_folder_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        select_folder_btn.setFixedSize(190, 45)
        select_folder_btn.clicked.connect(self.select_folder)

        clear_btn = PushButton("清空列表")
        clear_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        clear_btn.setFixedSize(190, 45)
        clear_btn.clicked.connect(self.clear_files)

        button_layout.addWidget(select_file_btn)
        button_layout.addWidget(select_folder_btn)
        button_layout.addWidget(clear_btn)
        button_layout.addStretch()
        file_card_layout.addLayout(button_layout)

        self.file_table = TableWidget()
        self.file_table.setColumnCount(1)
        self.file_table.setHorizontalHeaderLabels(["已选择文件"])
        self.file_table.horizontalHeader().setStretchLastSection(True)
        self.file_table.setMinimumHeight(180)
        self.file_table.setFont(QFont("Caribi, Microsoft YaHei", 10))  # 放大字体字号
        self.file_table.setStyleSheet("color: #000000;")
        file_card_layout.addWidget(self.file_table)

        main_layout.addWidget(file_card)

        settings_card = CardWidget()
        settings_card_layout = QVBoxLayout(settings_card)
        settings_card_layout.setContentsMargins(15, 15, 15, 15)
        settings_card_layout.setSpacing(12)

        settings_title = BodyLabel("操作和设置")
        settings_title.setFont(QFont("Caribi, Microsoft YaHei", 12, QFont.Bold))
        settings_title.setStyleSheet("color: #000000;")
        settings_card_layout.addWidget(settings_title)

        ops_layout = QHBoxLayout()
        ops_layout.setSpacing(10)

        process_btn = PrimaryPushButton("整理线索")
        process_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        process_btn.setFixedSize(190, 45)
        process_btn.clicked.connect(self.process_files)

        merge_btn = PrimaryPushButton("合并整理")
        merge_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        merge_btn.setFixedSize(190, 45)
        merge_btn.clicked.connect(self.merge_files)

        ops_layout.addWidget(process_btn)
        ops_layout.addWidget(merge_btn)
        ops_layout.addStretch()
        settings_card_layout.addLayout(ops_layout)

        params_layout = QVBoxLayout()
        params_layout.setSpacing(8)

        row1_layout = QHBoxLayout()
        row1_layout.setSpacing(10)

        output_label = BodyLabel("输出目录:")
        output_label.setFont(QFont("Caribi, Microsoft YaHei", 10))
        output_label.setStyleSheet("color: #000000;")
        self.output_line_edit = LineEdit()
        self.output_line_edit.setText(self.output_dir)
        self.output_line_edit.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.output_line_edit.setStyleSheet("color: #000000;")
        self.output_line_edit.setFixedHeight(40)

        browse_btn = PrimaryPushButton("选择输出目录")
        browse_btn.setFont(QFont("Caribi, Microsoft YaHei", 10))
        browse_btn.setFixedSize(190, 45)
        browse_btn.clicked.connect(self.browse_output_dir)

        row1_layout.addWidget(output_label)
        row1_layout.addWidget(self.output_line_edit, 1)
        row1_layout.addWidget(browse_btn)
        params_layout.addLayout(row1_layout)

        row2_layout = QHBoxLayout()
        row2_layout.setSpacing(10)

        anchor_label = BodyLabel("主播姓名:")
        anchor_label.setFont(QFont("Caribi, Microsoft YaHei", 10))
        anchor_label.setStyleSheet("color: #000000;")
        self.anchor_edit = LineEdit()
        self.anchor_edit.setText("晁天娇")
        self.anchor_edit.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.anchor_edit.setStyleSheet("color: #000000;")
        self.anchor_edit.setFixedHeight(40)

        local_label = BodyLabel("本地线索:")
        local_label.setFont(QFont("Caribi, Microsoft YaHei", 10))
        local_label.setStyleSheet("color: #000000;")
        self.local_edit = LineEdit()
        self.local_edit.setText("广东")
        self.local_edit.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.local_edit.setStyleSheet("color: #000000;")
        self.local_edit.setFixedHeight(40)

        row2_layout.addWidget(anchor_label)
        row2_layout.addWidget(self.anchor_edit, 1)
        row2_layout.addWidget(local_label)
        row2_layout.addWidget(self.local_edit, 1)
        params_layout.addLayout(row2_layout)

        settings_card_layout.addLayout(params_layout)

        options_layout = QHBoxLayout()
        options_layout.setSpacing(12)

        self.remove_empty_checkbox = CheckBox("删除空行")
        self.remove_empty_checkbox.setChecked(True)
        self.remove_empty_checkbox.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.remove_empty_checkbox.setStyleSheet("color: #000000;")

        self.remove_duplicates_checkbox = CheckBox("线索去重")
        self.remove_duplicates_checkbox.setChecked(True)
        self.remove_duplicates_checkbox.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.remove_duplicates_checkbox.setStyleSheet("color: #000000;")

        options_layout.addWidget(self.remove_empty_checkbox)
        options_layout.addWidget(self.remove_duplicates_checkbox)
        options_layout.addStretch()
        settings_card_layout.addLayout(options_layout)

        # 列选择区域
        column_label = BodyLabel("选择输出列:")
        column_label.setFont(QFont("Caribi, Microsoft YaHei", 10))
        column_label.setStyleSheet("color: #000000;")
        settings_card_layout.addWidget(column_label)

        # 创建列选择复选框的容器
        self.column_checkbox_widgets = {}
        columns_flow_layout = FlowLayout()
        columns_flow_layout.setSpacing(8)

        for col_name in self.column_checkboxes.keys():
            checkbox = CheckBox(col_name)
            checkbox.setChecked(self.column_checkboxes[col_name])
            checkbox.setFont(QFont("Caribi, Microsoft YaHei", 9))
            checkbox.setStyleSheet("color: #000000;")
            checkbox.stateChanged.connect(
                lambda state, name=col_name: self.on_column_checkbox_changed(name, state)
            )
            self.column_checkbox_widgets[col_name] = checkbox
            columns_flow_layout.addWidget(checkbox)

        settings_card_layout.addLayout(columns_flow_layout)

        main_layout.addWidget(settings_card)

        status_card = CardWidget()
        status_card_layout = QVBoxLayout(status_card)
        status_card_layout.setContentsMargins(15, 15, 15, 15)

        status_title = BodyLabel("日志")
        status_title.setFont(QFont("Caribi, Microsoft YaHei", 12, QFont.Bold))
        status_title.setStyleSheet("color: #000000;")
        status_card_layout.addWidget(status_title)

        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setPlaceholderText("处理日志将显示在这里...")
        self.status_text.setFont(QFont("Caribi, Microsoft YaHei", 10))
        self.status_text.setStyleSheet("color: #000000;")
        self.status_text.setMinimumHeight(120)
        status_card_layout.addWidget(self.status_text)

        main_layout.addWidget(status_card)

    def select_files(self):
        """选择Excel文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if files:
            self.add_files(files)

    def select_folder(self):
        """选择包含Excel文件的文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            excel_files = []
            # 遍历文件夹中的所有文件
            for root, dirs, files in os.walk(folder):
                for file in files:
                    if file.endswith((".xlsx", ".xls")):
                        excel_files.append(os.path.join(root, file))
            if excel_files:
                self.add_files(excel_files)
            else:
                self.status_text.append("所选文件夹中没有Excel文件")

    def add_files(self, files):
        """添加文件到列表"""
        try:
            for file in files:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    row_position = self.file_table.rowCount()
                    self.file_table.insertRow(row_position)
                    # 使用Text对QTableWidgetItem进行包装
                    item = QTableWidgetItem(file)
                    self.file_table.setItem(row_position, 0, item)
            self.status_text.append(f"已添加 {len(files)} 个文件")
        except Exception as e:
            self.status_text.append(f"添加文件时出错: {str(e)}")
            import traceback
            self.status_text.append(traceback.format_exc())

    def on_column_checkbox_changed(self, column_name, state):
        """处理列选择复选框变化"""
        self.column_checkboxes[column_name] = bool(state)

    def get_selected_columns(self):
        """获取用户选择的列列表"""
        selected_columns = [col for col, is_selected in self.column_checkboxes.items() if is_selected]
        # 如果选择了客户所在地，则自动添加是否本地线索列
        if "客户所在地" in selected_columns and "是否本地线索" not in selected_columns:
            selected_columns.insert(selected_columns.index("客户所在地") + 1, "是否本地线索")
        return selected_columns

    def clear_files(self):
        """清空文件列表"""
        self.selected_files.clear()
        self.file_table.setRowCount(0)
        self.status_text.append("文件列表已清空")

    def browse_output_dir(self):
        """浏览输出目录"""
        initial_dir = os.getcwd()
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录", initial_dir)
        if directory:
            try:
                # 测试目录是否有写入权限
                test_file = os.path.join(directory, ".test_permission.txt")
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
                self.output_dir = directory
                self.output_line_edit.setText(directory)
                self.status_text.append(f"已选择输出目录: {directory}")
            except PermissionError:
                self.status_text.append(f"权限提示: 需要写入权限才能使用目录 {directory}")
                self.status_text.append("请在系统设置中授予应用权限，或选择其他目录")
                self.status_text.append("建议选择: " + os.getcwd())

    def _process_single_file(self, df):
        """处理单个文件，提取所需列"""
        # 月份
        df["月份"] = datetime.now().strftime("%m月")
        
        # 更新时间
        df["更新时间"] = datetime.now().strftime("%H:%M")
        
        # 主播
        anchor_name = self.anchor_edit.text()
        df["主播"] = anchor_name
        
        # 客户所在地
        city_col = None
        for col in df.columns:
            if "城市" in col or "所在地区" in col:
                city_col = col
                break
        if city_col:
            df["客户所在地"] = df[city_col]
        else:
            df["客户所在地"] = ""
        
        # 是否本地线索
        df["是否本地线索"] = df["客户所在地"].apply(lambda x: "是" if "深圳" in str(x) else "否")
        
        # 跟进顾问
        df["跟进顾问"] = ""
        
        # 下发日期
        df["下发日期"] = datetime.now().strftime("%Y-%m-%d")
        
        # 客户姓名
        name_col = None
        for col in df.columns:
            if "用户信息" in col or "抖音昵称" in col or "昵称" in col:
                name_col = col
                break
        if name_col:
            df["客户姓名"] = df[name_col]
        else:
            df["客户姓名"] = ""
        
        # 抖音ID
        id_col = None
        id_keywords = ["抖音id", "抖音ID", "抖音账号", "客户抖音号"]
        for col in df.columns:
            col_lower = col.lower()
            for keyword in id_keywords:
                if keyword.lower() in col_lower:
                    id_col = col
                    break
            if id_col:
                break
        if id_col:
            df["抖音ID"] = df[id_col].apply(lambda x: str(x).strip() if pd.notna(x) else "")
        else:
            df["抖音ID"] = ""
        
        # 客户手机号
        phone_col = None
        phone_keywords = ["电话", "手机号", "客户手机号", "手机", "联系电话"]
        for col in df.columns:
            col_lower = col.lower()
            if "手机号" in col_lower or "phone" in col_lower:
                phone_col = col
                break
            for keyword in phone_keywords:
                if keyword.lower() in col_lower:
                    phone_col = col
                    break
            if phone_col:
                break
        if phone_col:
            df["客户手机号"] = df[phone_col].apply(self.clean_phone)
        else:
            df["客户手机号"] = ""
        
        # 微信号
        wechat_col = None
        wechat_keywords = ["微信号", "微信", "wechat"]
        for col in df.columns:
            col_lower = col.lower()
            if "微信号" in col_lower or "wechat" in col_lower:
                wechat_col = col
                break
            for keyword in wechat_keywords:
                if keyword.lower() in col_lower:
                    wechat_col = col
                    break
            if wechat_col:
                break
        if wechat_col:
            df["微信号"] = df[wechat_col].apply(self.clean_wechat)
        else:
            df["微信号"] = ""
        
        return df
    
    def _deduplicate_single_file(self, df, filename):
        """单文件去重，基于客户手机号和微信号"""
        original_len = len(df)
        # 基于客户手机号和微信号去重，保留第一条记录
        df = df.drop_duplicates(subset=["客户手机号", "微信号"], keep="first")
        self.status_text.append(f"文件 {filename} 移除了 {original_len - len(df)} 条重复记录，保留{len(df)}条有效记录！")
        return df
    
    def _deduplicate_after_merge(self, merged_df):
        """合并后去重，基于客户姓名，合并客户手机号和微信号"""
        original_len = len(merged_df)
        
        def merge_records(group):
            """合并重复记录"""
            if len(group) == 1:
                return group.iloc[0]
            
            # 初始化合并后的记录
            merged_record = group.iloc[0].copy()
            
            # 遍历所有记录，合并客户手机号和微信号
            for i in range(1, len(group)):
                current_record = group.iloc[i]
                
                # 合并客户手机号
                if pd.notna(current_record['客户手机号']) and current_record['客户手机号'] != '':
                    if pd.isna(merged_record['客户手机号']) or merged_record['客户手机号'] == '':
                        merged_record['客户手机号'] = current_record['客户手机号']
                    elif current_record['客户手机号'] != merged_record['客户手机号']:
                        merged_record['客户手机号'] = str(merged_record['客户手机号']) + ',' + str(current_record['客户手机号'])
                
                # 合并微信号
                if pd.notna(current_record['微信号']) and current_record['微信号'] != '':
                    if pd.isna(merged_record['微信号']) or merged_record['微信号'] == '':
                        merged_record['微信号'] = current_record['微信号']
                    elif current_record['微信号'] != merged_record['微信号']:
                        merged_record['微信号'] = str(merged_record['微信号']) + ',' + str(current_record['微信号'])
            
            return merged_record
        
        # 根据用户选择的列确定分组键
        if "客户姓名/抖音ID" in self.get_selected_columns():
            # 如果选择了客户姓名/抖音ID列，则基于该列分组
            grouped = merged_df.groupby(['客户姓名/抖音ID'])
        else:
            # 否则基于客户姓名分组
            grouped = merged_df.groupby(['客户姓名'])
        
        result_frames = []
        for name, group in grouped:
            merged_record = merge_records(group)
            if isinstance(merged_record, pd.DataFrame):
                result_frames.append(merged_record)
            else:
                result_frames.append(merged_record.to_frame().T)
        
        if result_frames:
            merged_df = pd.concat(result_frames, ignore_index=True)
        
        self.status_text.append(f"合并后移除了 {original_len - len(merged_df)} 条重复记录，保留{len(merged_df)}条有效记录！")
        return merged_df
    
    def process_files(self):
        """整理线索文件"""
        if not self.selected_files:
            self.status_text.append("请先选择要处理的Excel文件")
            return

        self.status_text.append("开始处理...")

        try:
            output_dir = self.output_dir

            # 创建输出目录（如果不存在）
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # 测试输出目录是否有写入权限
            test_file = os.path.join(output_dir, ".test_permission.txt")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)

            # 生成时间戳
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

            for file in self.selected_files:
                self.status_text.append(f"正在处理文件: {os.path.basename(file)}")
                # 读取Excel文件
                df = pd.read_excel(file)
                
                # 处理单个文件
                df = self._process_single_file(df)
                
                # 删除空行
                if self.remove_empty_checkbox.isChecked():
                    df = df.dropna(how='all')
                
                # 如果选择去重，则进行单文件去重
                if self.remove_duplicates_checkbox.isChecked():
                    df = self._deduplicate_single_file(df, os.path.basename(file))
                
                # 生成客户姓名/抖音ID列（如果用户选择了该列）
                if "客户姓名/抖音ID" in self.get_selected_columns():
                    def get_customer_name(row):
                        try:
                            if "抖音ID" in row and row["抖音ID"] and row["抖音ID"] != "":
                                return row["抖音ID"]
                            elif "客户姓名" in row and row["客户姓名"] and row["客户姓名"] != "":
                                return f"{row['客户姓名']}(抖音昵称)"
                            else:
                                return ""
                        except Exception:
                            return ""
                    
                    df["客户姓名/抖音ID"] = df.apply(get_customer_name, axis=1)
                
                # 保存处理后的文件
                output_file = os.path.join(output_dir, f"整理线索_{timestamp}.xlsx")
                # 按照要求的顺序排列列并保存
                required_columns = self.get_selected_columns()
                df[required_columns].to_excel(output_file, index=False)
                self.status_text.append(f"已保存: {output_file}")
                
                # 将处理过的源文件路径添加到processed_files中
                if file not in self.processed_files:
                    self.processed_files.append(file)

            self.status_text.append("文件已处理完成！")
            
            # 自动清空文件列表
            self.clear_files()
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("完成")
# 使用富文本装饰文字
            status_title = "<h3 style='color: #2E7D32;'>✅ 处理成功</h3>"
            path_info = f"<p style='color: #555;'>文件已保存至：</p><code style='color: #005FB8;'>{output_dir}</code>"
            ask_text = "<p><br><b>您想立即查看生成的文件吗？</b></p>"
                
            msg_box.setText(status_title + "<hr>" + path_info + ask_text)
             # 修改按钮为直观的中文文案
            open_btn = msg_box.addButton("打开", QMessageBox.YesRole)
            close_btn = msg_box.addButton("不打开", QMessageBox.NoRole)
            msg_box.setDefaultButton(open_btn)
            msg_box.exec_()   
                # 判断点击的是否为“立即打开”按钮
            if msg_box.clickedButton() == open_btn:
                os.startfile(output_file)

        except Exception as e:
            self.status_text.append(f"错误: {str(e)}")
            QMessageBox.critical(self, "错误", f"处理文件时出错: {str(e)}")

    def merge_files(self):
        """合并整理线索文件"""
        if not self.selected_files:
            self.status_text.append("请先选择Excel文件")
            return

        self.status_text.append("开始合并文件...")

        try:
            output_dir = self.output_dir

            # 创建输出目录（如果不存在）
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            # 测试输出目录是否有写入权限
            test_file = os.path.join(output_dir, ".test_permission.txt")
            with open(test_file, "w") as f:
                f.write("test")
            os.remove(test_file)

            all_data = []

            for file in self.selected_files:
                self.status_text.append(f"读取文件: {os.path.basename(file)}")
                # 读取Excel文件
                df = pd.read_excel(file)
                
                # 处理单个文件
                df = self._process_single_file(df)
                
                # 删除空行
                if self.remove_empty_checkbox.isChecked():
                    df = df.dropna(how='all')
                
                # 如果选择去重，则进行单文件去重
                if self.remove_duplicates_checkbox.isChecked():
                    df = self._deduplicate_single_file(df, os.path.basename(file))
                
                # 生成客户姓名/抖音ID列（如果用户选择了该列）
                if "客户姓名/抖音ID" in self.get_selected_columns():
                    def get_customer_name(row):
                        try:
                            if "抖音ID" in row and row["抖音ID"] and row["抖音ID"] != "":
                                return row["抖音ID"]
                            elif "客户姓名" in row and row["客户姓名"] and row["客户姓名"] != "":
                                return f"{row['客户姓名']}(抖音昵称)"
                            else:
                                return ""
                        except Exception:
                            return ""
                    
                    df["客户姓名/抖音ID"] = df.apply(get_customer_name, axis=1)
                
                all_data.append(df)
                
                # 将处理过的源文件路径添加到processed_files中
                if file not in self.processed_files:
                    self.processed_files.append(file)

            if all_data:
                # 合并所有文件
                merged_df = pd.concat(all_data, ignore_index=True)

                # 如果选择去重，则进行合并后去重
                if self.remove_duplicates_checkbox.isChecked():
                    merged_df = self._deduplicate_after_merge(merged_df)

                # 删除空行
                if self.remove_empty_checkbox.isChecked():
                    merged_df = merged_df.dropna(how='all')

                # 生成客户姓名/抖音ID列（如果用户选择了该列）
                if "客户姓名/抖音ID" in self.get_selected_columns():
                    def get_customer_name(row):
                        try:
                            if "抖音ID" in row and row["抖音ID"] and row["抖音ID"] != "":
                                return row["抖音ID"]
                            elif "客户姓名" in row and row["客户姓名"] and row["客户姓名"] != "":
                                return f"{row['客户姓名']}(抖音昵称)"
                            else:
                                return ""
                        except Exception:
                            return ""
                    
                    merged_df["客户姓名/抖音ID"] = merged_df.apply(get_customer_name, axis=1)

                # 生成时间戳
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                # 保存合并后的文件
                output_file = os.path.join(output_dir, f"合并线索_{timestamp}.xlsx")
                # 按照要求的顺序排列列并保存
                required_columns = self.get_selected_columns()
                merged_df[required_columns].to_excel(output_file, index=False)
                self.status_text.append(f"已保存: {output_file}")
                self.status_text.append(f"合并完成！共 {len(merged_df)} 条记录")
                
                # 自动清空文件列表
                self.clear_files()
                
                # 弹窗提示文件保存位置和是否打开文件
                msg_box = QMessageBox()
                msg_box.setWindowTitle("完成")
                msg_box.setIcon(QMessageBox.Information)

                # 使用富文本装饰文字
                status_title = "<h3 style='color: #2E7D32;'>✅ 处理成功</h3>"
                path_info = f"<p style='color: #555;'>文件已保存至：</p><code style='color: #005FB8;'>{output_dir}</code>"
                ask_text = "<p><br><b>您想立即查看生成的文件吗？</b></p>"
                
                msg_box.setText(status_title + "<hr>" + path_info + ask_text)

                # 修改按钮为直观的中文文案
                open_btn = msg_box.addButton("立即打开", QMessageBox.YesRole)
                close_btn = msg_box.addButton("暂不打开", QMessageBox.NoRole)
                msg_box.setDefaultButton(open_btn)

                msg_box.exec_()
                
                if msg_box.clickedButton() == open_btn:
                    os.startfile(output_file)

        except Exception as e:
            self.status_text.append(f"错误: {str(e)}")
            QMessageBox.critical(self, "错误", f"合并文件时出错: {str(e)}")

    def clean_phone(self, phone):
        """清理手机号"""
        if pd.isna(phone):
            return ""
        # 处理浮点数转换为字符串时可能出现的.0问题
        if isinstance(phone, float):
            # 检查是否为整数
            if phone.is_integer():
                phone = int(phone)
        phone_str = str(phone).strip()
        # 提取数字
        digits = ''.join(filter(str.isdigit, phone_str))
        return digits if digits else ""

    def clean_wechat(self, wechat):
        """清理微信号"""
        if pd.isna(wechat):
            return ""
        wechat_str = str(wechat).strip()
        return wechat_str if wechat_str else ""
    
    def closeEvent(self, event):
        """窗口关闭事件"""
        # 如果有处理过的源文件，询问是否删除
        if self.processed_files:
            msg_box = QMessageBox()
            msg_box.setWindowTitle("退出确认")
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setText("是否删除刚才处理的源文件？")
            
            
            # 添加按钮
            yes_btn = msg_box.addButton("是", QMessageBox.YesRole)
            no_btn = msg_box.addButton("否", QMessageBox.NoRole)
            cancel_btn = msg_box.addButton("取消", QMessageBox.RejectRole)
            
            msg_box.setDefaultButton(yes_btn)
            msg_box.exec_()
            
            if msg_box.clickedButton() == yes_btn:
                # 删除处理过的源文件
                deleted_count = 0
                for file in self.processed_files:
                    try:
                        if os.path.exists(file):
                            os.remove(file)
                            deleted_count += 1
                    except Exception as e:
                        print(f"删除文件失败: {file}, 错误: {str(e)}")
                
                if deleted_count > 0:
                    self.status_text.append(f"已删除 {deleted_count} 个处理过的源文件")
                event.accept()
            elif msg_box.clickedButton() == no_btn:
                # 不删除源文件，直接退出
                event.accept()
            else:
                # 取消退出
                event.ignore()
        else:
            # 没有处理过的源文件，直接退出
            event.accept()


if __name__ == "__main__":
    """主函数"""
    app = QApplication(sys.argv)
    window = ExcelTool()
    window.show()
    sys.exit(app.exec_())
