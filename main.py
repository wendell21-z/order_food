import sys

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QBrush
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QPushButton, QLabel, QGroupBox,
                               QMessageBox, QFileDialog, QTabWidget, QTableWidget,
                               QTableWidgetItem, QSplitter)


class NumericTableWidgetItem(QTableWidgetItem):
    """Custom QTableWidgetItem that sorts based on numeric value."""

    def __init__(self, value):
        super().__init__(str(value))
        self.value = value  # Store numeric value for sorting

    def __lt__(self, other):
        # Ensure we only compare with same type
        if isinstance(other, NumericTableWidgetItem):
            return self.value < other.value
        # Fallback to default behavior
        return super().__lt__(other)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("订餐支付校验")
        # self.setGeometry(100, 100, 1200, 700)

        # 存储两个Excel文件路径
        self.order_excel_path = None
        self.file_path_2 = None

        # 存储处理结果
        self.order_map = {}
        self.error_bills = []

        self.statistics = {}
        """
        {
            ('焊装车间','M9'):{
                dish:{
                    '麻花':{
                        'amount':2,
                        'price':7
                    },
                    'Q蛋肠':{
                        'amount':1,
                        'price':3.5
                    }
                },
                price:200
            }
        }
        """

        self.dish_count = {}
        """
        {
            "Q蛋肠":10,
        }
        """

        self.total_price = 0

        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        excel_widget = QWidget()
        excel_layout = QHBoxLayout(excel_widget)

        # 创建两个上传组件组
        self.group_box_1 = self.create_upload_group("订单表", 1)
        self.group_box_2 = self.create_upload_group("支付表", 2)

        excel_layout.addWidget(self.group_box_1)
        excel_layout.addWidget(self.group_box_2)
        main_layout.addWidget(excel_widget)

        # 创建主显示区域 - 使用选项卡
        main_tabs = QTabWidget()

        # ============  第一个选项卡：菜品统计 ============
        self.statistics_widget(main_layout, main_tabs)

        # ============ 第二个选项卡：订单详情 ============
        self.order_detail_widget(main_layout, main_tabs)

    def statistics_widget(self, main_layout: QVBoxLayout, main_tabs: QTabWidget):
        widget1 = QWidget()
        layout1 = QVBoxLayout(widget1)

        # 统计按钮,如果订单表没上传，按钮禁用
        process_button = QPushButton("统计")
        process_button.clicked.connect(self.compute_statistics)
        layout1.addWidget(process_button)

        widget2 = QWidget()
        layout2 = QHBoxLayout(widget2)
        layout1.addWidget(widget2)

        dish_count_group = QGroupBox("菜品数量统计")
        dish_count_layout = QVBoxLayout(dish_count_group)
        layout2.addWidget(dish_count_group)
        # 表1 显示各个dish 的数量
        self.dish_count_table = QTableWidget()
        self.dish_count_table.setSortingEnabled(True)
        self.dish_count_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dish_count_table.setColumnCount(2)
        self.dish_count_table.setHorizontalHeaderLabels(['菜品名称', '数量'])
        dish_count_layout.addWidget(self.dish_count_table)

        dish_group = QGroupBox("菜品统计")
        dish_layout = QVBoxLayout(dish_group)
        layout2.addWidget(dish_group)
        # 表2 显示各个送货点和总金额
        self.statistics_table = QTableWidget()
        self.statistics_table.setSortingEnabled(True)
        self.statistics_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.statistics_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.statistics_table.setColumnCount(3)
        self.statistics_table.setHorizontalHeaderLabels(['车间', '送货点', '总金额'])
        self.statistics_table.itemSelectionChanged.connect(self.on_statistics_table_selected_change)
        dish_layout.addWidget(self.statistics_table)

        # 表3 显示送货点对应的dish_name,dish_price,dish_amount
        self.point_statistics_table = QTableWidget()
        self.point_statistics_table.setSortingEnabled(True)
        self.point_statistics_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.point_statistics_table.setColumnCount(2)
        self.point_statistics_table.setHorizontalHeaderLabels(['菜品名称', '数量'])
        dish_layout.addWidget(self.point_statistics_table)

        # 导出为excel按钮
        export_button = QPushButton("导出为excel")
        export_button.clicked.connect(self.export_statistics_table)
        layout1.addWidget(export_button)

        main_tabs.addTab(widget1, "菜品统计")
        main_layout.addWidget(main_tabs)

    def order_detail_widget(self, main_layout, main_tabs):
        widget1 = QWidget()
        layout1 = QVBoxLayout(widget1)
        main_layout.addWidget(widget1)

        # 校验按钮
        process_button = QPushButton("校验")
        process_button.clicked.connect(self.compute_pay)
        layout1.addWidget(process_button)

        widget2 = QTabWidget()
        layout1.addWidget(widget2)

        # tab1 - 订单信息
        detail_widget = QWidget()
        order_detail_layout = QHBoxLayout(detail_widget)

        # 水平拆分器 - 左侧订单列表，右侧详情
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：订单信息表格
        self.order_table = QTableWidget()
        self.order_table.setSortingEnabled(True)  # 开启排序功能
        self.order_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 禁止编辑
        self.order_table.setColumnCount(7)  # 工号、姓名、车间-工段、取餐点、支付状态、金额差值、余额
        self.order_table.setHorizontalHeaderLabels(
            ['工号', '姓名', '车间-工段', '取餐点', '支付状态', '金额差值', '余额'])
        self.order_table.setSelectionBehavior(QTableWidget.SelectRows)  # 设置为行选择
        self.order_table.itemSelectionChanged.connect(self.on_order_selection_changed)  # 连接选择变化事件

        # 将订单表格放在滚动容器中
        order_group = QGroupBox("订单信息")

        # 创建一个垂直布局来包含按钮和表格
        order_layout = QVBoxLayout(order_group)

        # 创建一个水平布局来包含导出按钮
        button_layout = QHBoxLayout()
        export_button = QPushButton("导出Excel")
        export_button.setStyleSheet("font-size: 12px; padding: 5px; width:80px;")
        export_button.clicked.connect(self.export_order_table)
        button_layout.addWidget(export_button)
        button_layout.addStretch()  # 添加弹性空间，使按钮靠左对齐

        order_layout.addWidget(self.order_table)
        order_layout.addLayout(button_layout)

        # 右侧：详情显示区域 - 垂直拆分
        right_splitter = QSplitter(Qt.Vertical)

        # 菜品信息表格
        self.dish_table = QTableWidget()
        self.dish_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dish_table.setColumnCount(3)  # 菜名、数量、单价
        self.dish_table.setHorizontalHeaderLabels(['菜名', '数量', '单价'])
        dish_group = QGroupBox("菜品详情")
        dish_layout = QVBoxLayout(dish_group)
        dish_layout.addWidget(self.dish_table)

        # 账单信息表格
        self.bill_table = QTableWidget()
        self.bill_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bill_table.setColumnCount(4)  # 消费金额、余额、发生时间、POS机
        self.bill_table.setHorizontalHeaderLabels(['消费金额', '余额', '发生时间','POS机'])
        bill_group = QGroupBox("账单详情")
        bill_layout = QVBoxLayout(bill_group)
        bill_layout.addWidget(self.bill_table)

        right_splitter.addWidget(dish_group)
        right_splitter.addWidget(bill_group)
        right_splitter.setSizes([300, 300])  # 设置初始大小比例

        # 将左右两部分添加到主拆分器
        splitter.addWidget(order_group)
        splitter.addWidget(right_splitter)
        splitter.setSizes([750, 350])  # 设置左右比例

        order_detail_layout.addWidget(splitter)
        widget2.addTab(detail_widget, "订单详情")

        # tab2 - 无订单支付
        error_widget = QWidget()
        error_layout = QVBoxLayout(error_widget)

        # 无订单支付表格
        self.error_table = QTableWidget()
        self.error_table.setSortingEnabled(True)  # 开启排序功能
        self.error_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 禁止编辑
        self.error_table.setColumnCount(6)  # 工号、姓名、发生额、发生后库余额、支付时间、POS机
        self.error_table.setHorizontalHeaderLabels(['工号', '姓名', '支付金额', '支付后余额', '支付时间', 'POS机'])
        error_group = QGroupBox("无订单支付")
        error_group_layout = QVBoxLayout(error_group)
        error_group_layout.addWidget(self.error_table)
        error_layout.addWidget(error_group)
        widget2.addTab(error_widget, "无订单支付")

        main_tabs.addTab(widget1, "订单校验")

    def create_upload_group(self, title, group_number):
        """创建一个包含上传组件的组 - 垂直结构"""
        group = QGroupBox(title)
        layout = QHBoxLayout(group)

        # 文件路径显示标签
        file_label = QLabel("未选择文件")
        file_label.setWordWrap(True)

        # 上传按钮
        upload_button = QPushButton(f"上传{title}")
        upload_button.clicked.connect(
            lambda: self.select_excel_file(group_number)
        )
        upload_button.setFixedWidth(100)

        layout.addWidget(file_label)
        layout.addWidget(upload_button)

        # 保存标签引用以便后续更新
        if group_number == 1:
            self.order_file_label = file_label
            # 清空相关数据
            self.order_map = {}
            # 更新相关table

        else:
            self.file_label_2 = file_label

        return group

    def select_excel_file(self, group_number):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            "",
            "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            if group_number == 1:
                self.order_excel_path = file_path
                self.order_file_label.setText(f"已选择: {file_path.split('/')[-1]}")
            else:
                self.file_path_2 = file_path
                self.file_label_2.setText(f"已选择: {file_path.split('/')[-1]}")

    def compute_pay(self):
        """处理两个Excel文件"""
        if not self.order_excel_path or not self.file_path_2:
            QMessageBox.warning(self, "警告", "请先选择两个Excel文件！")
            return

        # 清空旧值
        self.error_bills = []  # 清空旧值

        try:
            # 导入必要的类和函数
            from compute import create_order, add_bill_info

            # 处理文件
            self.compute_statistics()
            self.error_bills = add_bill_info(self.order_map, self.file_path_2)

            # 计算每个订单的差额和余额
            for work_no, order in self.order_map.items():
                order.calc_amount_diff()
                order.calc_balance()

            # 更新表格显示
            self.update_order_table()
            self.update_error_table()

            QMessageBox.information(self, "处理完成",
                                    f"成功处理文件！\n订单数量: {len(self.order_map)}\n错误账单数量: {len(self.error_bills)}")

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理文件时出错:\n{str(e)}")


    def update_order_table(self):
        """更新订单信息表格"""
        if not self.order_map:
            self.order_table.setRowCount(0)
            return

        self.order_table.setRowCount(len(self.order_map))

        for row, (work_no, order) in enumerate(self.order_map.items()):
            self.order_table.setItem(row, 0, NumericTableWidgetItem(order.work_no))
            self.order_table.setItem(row, 1, QTableWidgetItem(order.name))
            self.order_table.setItem(row, 2, QTableWidgetItem(order.workshop_section))
            self.order_table.setItem(row, 3, QTableWidgetItem(order.take_out_point))

            pay_status_item = QTableWidgetItem(order.pay_status)
            pay_status_item.setForeground(
                QBrush(QColor(0, 128, 0) if order.pay_status == '正常' else QColor(255, 0, 0)))
            self.order_table.setItem(row, 4, pay_status_item)

            self.order_table.setItem(row, 5, NumericTableWidgetItem(order.amount_diff))
            self.order_table.setItem(row, 6, NumericTableWidgetItem(order.balance))

        # 调整列宽
        self.order_table.resizeColumnsToContents()

    def update_dish_table(self, order):
        """更新菜品信息表格"""
        if not order.dishes:
            self.dish_table.setRowCount(0)
            return

        self.dish_table.setRowCount(len(order.dishes))

        for row, dish in enumerate(order.dishes):
            self.dish_table.setItem(row, 0, QTableWidgetItem(dish.name))
            self.dish_table.setItem(row, 1, NumericTableWidgetItem(dish.num))
            self.dish_table.setItem(row, 2, NumericTableWidgetItem(dish.price))

        # 调整列宽
        self.dish_table.resizeColumnsToContents()

    def update_statistics_table(self):
        self.statistics_table.setRowCount(len(self.statistics))
        for row, (workshop, point) in enumerate(self.statistics.keys()):
            self.statistics_table.setItem(row, 0, QTableWidgetItem(workshop))
            self.statistics_table.setItem(row, 1, QTableWidgetItem(point))
            self.statistics_table.setItem(row, 2, NumericTableWidgetItem(self.statistics[workshop, point]['price']))

    # 菜品数据量统计表
    def update_dish_count_table(self):
        self.dish_count_table.setRowCount(len(self.dish_count))
        for row, (dish_name, count) in enumerate(self.dish_count.items()):
            self.dish_count_table.setItem(row, 0, QTableWidgetItem(dish_name))
            self.dish_count_table.setItem(row, 1, NumericTableWidgetItem(count))

    # 订单详情表
    def update_bill_table(self, order):
        """更新账单信息表格"""
        if not order.bills:
            self.bill_table.setRowCount(0)
            return

        self.bill_table.setRowCount(len(order.bills))

        for row, bill in enumerate(order.bills):
            self.bill_table.setItem(row, 0, NumericTableWidgetItem(bill.amount))
            self.bill_table.setItem(row, 1, NumericTableWidgetItem(bill.balance))
            self.bill_table.setItem(row, 2, QTableWidgetItem(str(bill.time)))
            self.bill_table.setItem(row, 2, QTableWidgetItem(str(bill.pos)))

        # 调整列宽
        self.bill_table.resizeColumnsToContents()

    # 无订单支付表
    def update_error_table(self):
        """更新错误账单表格"""
        if not self.error_bills:
            self.error_table.setRowCount(0)
            return

        self.error_table.setRowCount(len(self.error_bills))

        for row, error_bill in enumerate(self.error_bills):
            self.error_table.setItem(row, 0, NumericTableWidgetItem(error_bill.work_no))
            self.error_table.setItem(row, 1, QTableWidgetItem(error_bill.name))
            self.error_table.setItem(row, 2, NumericTableWidgetItem(error_bill.amount))
            self.error_table.setItem(row, 3, NumericTableWidgetItem(error_bill.balance))
            self.error_table.setItem(row, 4, QTableWidgetItem(str(error_bill.time)))
            self.error_table.setItem(row, 5, QTableWidgetItem(error_bill.pos))

    def on_order_selection_changed(self):
        """当订单表格选中行变化时，更新右侧的dishes和bills表格"""
        selected_rows = self.order_table.selectedItems()
        if not selected_rows:
            # 清空右侧表格
            self.dish_table.setRowCount(0)
            self.bill_table.setRowCount(0)
            return

        # 获取选中行的第一个单元格（工号）
        row = selected_rows[0].row()
        work_no_item = self.order_table.item(row, 0)
        if work_no_item:
            work_no = int(work_no_item.text())

            # 根据工号找到对应的订单
            if work_no in self.order_map:
                order = self.order_map[work_no]

                # 更新菜品表格
                self.update_dish_table(order)

                # 更新账单表格
                self.update_bill_table(order)

    def export_order_table(self):
        """导出订单表格到Excel文件"""
        if not self.order_map:
            QMessageBox.warning(self, "警告", "没有数据可导出！")
            return

        # 选择保存文件路径
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存Excel文件",
            "订单信息.xlsx",
            "Excel Files (*.xlsx *.xls)"
        )

        if file_path:
            try:
                # 创建DataFrame
                data = []
                for work_no, order in self.order_map.items():
                    row = [
                        order.work_no,
                        order.name,
                        order.workshop_section,
                        order.take_out_point,
                        order.pay_status,
                        order.amount_diff,
                        order.balance,
                        '；'.join([f'{dish.name}<{dish.num}>' for dish in order.dishes])
                    ]
                    data.append(row)

                # 获取表头
                headers = ['工号', '姓名', '车间-工段', '取餐点', '支付状态', '金额差值', '余额', '菜品(数量)']
                df = pd.DataFrame(data, columns=headers)

                # 保存到Excel
                df.to_excel(file_path, index=False)

                QMessageBox.information(self, "导出成功", f"订单信息已成功导出到:\n{file_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出文件时出错:\n{str(e)}")

    def export_statistics_table(self):
        # todo
        """导出统计表格到Excel文件"""
        if not self.statistics:
            QMessageBox.warning(self, "警告", "没有数据可导出！")
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存Excel文件",
            "统计信息.xlsx",
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            df = pd.DataFrame()
            header = ['车间', '取餐点'] + list(self.dish_count.keys())
            row_n = 0
            for key, value in self.statistics.items():
                row_n += 1
                workshop = key[0]
                point = key[1]
                df.at[row_n, '车间'] = workshop
                df.at[row_n, '取餐点'] = point
                for dish_name in header[2:]:
                    dish_num = 0
                    if dish_name in value['dish']:
                        dish_num = value['dish'][dish_name]['amount']
                    df.at[row_n, dish_name] = dish_num

            df.to_excel(file_path, index=True, header=header)
            QMessageBox.information(self, "导出成功", f"统计信息已成功导出到:\n{file_path}")

    def compute_statistics(self):
        # 如果订单表为空，提示用户需要上传
        if not self.order_excel_path:
            QMessageBox.warning(self, "警告", "请先上传订单信息！")
            return

        self.order_map = {}  # 清空旧值

        # 导入必要的类和函数
        from compute import create_order
        self.order_map = create_order(self.order_excel_path)

        """
        计算订单统计信息
        :return:
        """
        # 按车间和取餐点统计餐品种类和数量
        grouped = {}
        for order in self.order_map.values():
            key = (order.workshop, order.take_out_point)
            for dish in order.dishes:
                grouped.setdefault(key, []).append(dish)

        statistics = {}
        """
        target = 
        {
            (work_shop, take_out_point): {
                dish:{
                    dish_name: [count,total_price],
                },
                price:0,                
            }
        }
        """
        total_price = 0
        total_count = {}
        """
        {
            dish_name: count,
        }
        """
        for key, dishes in grouped.items():
            dish_map = {}
            """
            dish_map = 
            {
                dish_name: [count,total_price],
            }
            """
            dish_total_price = 0
            for dish in dishes:
                dish_value = dish_map.setdefault(dish.name, {'amount': 0, 'price': 0.0})
                dish_value['amount'] += dish.num
                price = dish.price * dish.num
                dish_value['price'] += price
                dish_total_price += price
                total_price += price
                total_count[dish.name] = total_count.get(dish.name, 0) + dish.num

            sv = statistics.setdefault(key, {})
            sv['dish'] = dish_map
            sv['price'] = dish_total_price
        self.statistics = statistics
        self.total_price = total_price
        self.dish_count = total_count

        self.update_statistics_table()
        self.update_dish_count_table()

    def on_statistics_table_selected_change(self):
        select_rows = self.statistics_table.selectedItems()
        if not select_rows:
            self.point_statistics_table.setRowCount(0)
            return
        row = select_rows[0].row()
        workshop, point = self.statistics_table.item(row, 0).text(), self.statistics_table.item(row, 1).text()
        dishes = self.statistics[workshop, point]['dish']
        self.point_statistics_table.setRowCount(len(dishes))
        for row, (dish_name, dish_value) in enumerate(dishes.items()):
            self.point_statistics_table.setItem(row, 0, QTableWidgetItem(dish_name))
            self.point_statistics_table.setItem(row, 1, NumericTableWidgetItem(dish_value['amount']))


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
