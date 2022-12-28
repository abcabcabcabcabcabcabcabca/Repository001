import sys
from UI_Driver import *
from HTMLObjectScanner import *
import threading
import time
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PageGenerator import *


class EmptyDelegate(QItemDelegate):
    def __init__(self, parent):
        super(EmptyDelegate, self).__init__(parent)

    def createEditor(self, parent, option, index):
        pass

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.DisplayRole)
        editor.setText(str(value))

    def setModelData(self, editor, model, index):
        pass


class MyWindow(QMainWindow, Ui_HTMLScanner):
    def __init__(self, parent=None):
        super(MyWindow, self).__init__(parent)
        self.setupUi(self)
        self.center()
        self.setMaximumSize(1847, 836)
        self.setMinimumSize(1847, 836)
        self.ScanButton.pressed.connect(self.scan_browsers_note)

        self.left_sm = QStandardItemModel()
        self.right_sm = QStandardItemModel()
        self.urllistView_sm = QStandardItemModel()

        self.urllistView.horizontalHeader().setHighlightSections(False)
        self.urllistView.horizontalHeader().setFont(QFont("Arial", 10, QFont.Bold))
        self.urllistView.verticalHeader().hide()
        self.urllistView.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.urllistView.setAlternatingRowColors(True)
        self.urllistView.setShowGrid(True)
        self.urllistView.horizontalHeader().setStyleSheet("QHeaderView::section{background-color: rgb(238,232,205)};")
        self.urllistView.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.urllistView.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.urllistView.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)

        self.urllistView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.urllistView.customContextMenuRequested.connect(self.url_text_right_menu)

        self.all_loaded_urls = []
        self.loaded_data_in_left = []
        self.loaded_data_in_right = []
        self.slm = QStringListModel()
        self.tag_list = []
        self.selected_url = ''
        self.selected_url_date = ''
        self.frame_data = ''
        self.showing_data = ''
        self.arg_details = ''
        self.validate_details = ''
        self.full_screen_name = ''
        self.coordinates = []
        self.coordinates_image = ''
        self.list_header_order = {}
        self.full_scan_flag = False
        tab_style = "QTabBar::tab{background-color:rbg(255,255,255,0);}" + \
                    "QTabBar::tab:selected{color:red;background-color:rbg(255,200,255);} "
        self.tabWidget.setStyleSheet(tab_style)
        self.class_name = ''
        self.left_filter = QSortFilterProxyModel()
        self.left_filter.setDynamicSortFilter(True)

        self.lefttableView.horizontalHeader().setHighlightSections(False)
        self.lefttableView.horizontalHeader().setFont(QFont("QFont", 10, QFont.Bold))
        self.lefttableView.verticalHeader().hide()
        self.lefttableView.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.lefttableView.setAlternatingRowColors(True)
        self.lefttableView.setShowGrid(True)
        self.lefttableView.horizontalHeader().setStyleSheet("QHeaderView::section{background-color: rgb(238,232,205)};")
        self.lefttableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.lefttableView.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.lefttableView.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.righttableView.horizontalHeader().setHighlightSections(False)
        self.righttableView.horizontalHeader().setFont(QFont("QFont", 10, QFont.Bold))
        self.righttableView.verticalHeader().hide()
        self.righttableView.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.righttableView.setAlternatingRowColors(True)
        self.righttableView.setItemDelegateForColumn(0, EmptyDelegate(self))
        self.righttableView.setItemDelegateForColumn(2, EmptyDelegate(self))
        self.righttableView.horizontalHeader().setStyleSheet("QHeaderView::section{background-color:rgb(238,232,205)};")
        self.righttableView.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.righttableView.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.load_url_data(URL_PATH)

        self.scanner = FetchHtmlInfo()

        self.lefttableView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.lefttableView.customContextMenuRequested.connect(self.custom_right_menu)
        self.righttableView.setContextMenuPolicy(Qt.CustomContextMenu)
        self.righttableView.customContextMenuRequested.connect(self.right_custom_right_menu)

        self.tabWidget.currentChanged.connect(self.slot_small_tab)

        self.searchcomboBox.addItem("Tag")
        self.searchcomboBox.addItem("Attributes")
        self.searchcomboBox.addItem("Text")
        self.searchcomboBox.setCurrentIndex(1)

        self.image_container.setText('')

    def validate_rule(self):
        if self.personal_rules.toPlainText().strip() == '':
            return
        else:
            time_delay = 0
            if self.time_delay.text() != '':
                time_delay = int(self.time_delay.text().strip())
            time.sleep(time_delay)
            runtime, url, data_path = self.scanner.validate_rule_by_driver(self.personal_rules.toPlainText().strip())
            self.full_screen_name = os.path.join(data_path, "full_screen_shot.png")
            if os.path.isfile(self.full_screen_name):
                self.image_container.setPixmap(QPixmap(self.full_screen_name))
                self.image_container.setScaledContents(True)

    def change_filter_flag(self):
        self.searchText.setText("")

    def url_text_right_menu(self):
        try:
            menu = QMenu(self)
            opt1 = menu.addAction("Delete")
            action = menu.exec_(QCursor.pos())
            if action == opt1:
                self.delete_url_and_data()

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def delete_url_and_data(self):
        try:
            indexes = self.urllistView.selectionModel().selectedIndexes()
            time_selected = self.urllistView_sm.item(indexes[0].row(), 0).text().\
                replace(':', '_').replace('-', '_').replace(' ', '_').strip()

            if time_selected != self.selected_url_date:
                folder_path = '{0}\\{1}'.format(EXCEL_FOLDER_PATH, time_selected)
                if os.path.exists(folder_path):
                    shutil.rmtree(folder_path)
                self.urllistView_sm.removeRow(self.urllistView.currentIndex().row())
                with open(URL_PATH, "w+") as file_new:
                    for row_index in range(self.urllistView_sm.rowCount()):
                        line = '{0},{1}\n'.format(self.urllistView_sm.item(row_index, 0).text().strip(),
                                                self.urllistView_sm.item(row_index, 1).text().strip())
                        file_new.write(line)
            else:
                QMessageBox.information(self, "Information", "You cannot delete the current opening Excel.")
        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def view_tag_images(self):
        try:
            tag_option = self.tagcomboBox.currentText().replace('<', '').replace('>', '')

            name = parse_url(self.selected_url)
            file_name = EXCEL_FOLDER_PATH + '\\{0}\\{1}\\{2}'.format(self.selected_url_date, name.strip(),
                                                                     "{0}_image.png".format(tag_option))

            if os.path.isfile(file_name):
                self.image_container.setPixmap(QPixmap(file_name))
                self.image_container.setScaledContents(True)

            if tag_option != 'ALL':
                self.left_filter.setFilterKeyColumn(1)
                self.left_filter.setFilterFixedString(tag_option)
            else:
                self.left_filter.setFilterKeyColumn(1)
                self.left_filter.setFilterFixedString("")

                if os.path.isfile(self.full_screen_name):
                    self.image_container.setPixmap(QPixmap(self.full_screen_name))
                    self.image_container.setScaledContents(True)

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def slot_small_tab(self):
        try:
            if self.tabWidget.currentIndex() == 2:
                self.textEdit.setText('')
                switch_flag = self.check_right_values() is False and self.check_dup_records()
                if switch_flag is True:
                    exp_data = pd.DataFrame(columns=['locator', 'name', 'frame'])
                    for row_index in range(self.right_sm.rowCount()):
                        exp_data = exp_data.append({'locator': self.right_sm.item(row_index, 0).text(),
                                                    'name': self.right_sm.item(row_index, 1).text(),
                                                    'frame': self.right_sm.item(row_index, 2).text()},
                                                   ignore_index=True)
                    exp_data.drop_duplicates(['locator'], keep='first', inplace=True)
                    exp_data_sort = exp_data.sort_values(by='frame', ascending=True)
                    code_generator = PageGenerator(exp_data_sort, frame_data=self.frame_data)
                    self.arg_details = code_generator.parse_html_data(self.selected_url)
                    self.textEdit.setPlainText(self.arg_details)
                else:
                    self.tabWidget.setCurrentIndex(1)
        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def check_right_values(self):
        try:
            check_flag = 'True'
            for right_index in range(self.right_sm.rowCount()):
                name_text = self.right_sm.item(right_index, 1).text().strip()
                if name_text != '' and (name_text[0].encode('UTF-8').isalpha() or name_text[0] == '_'):
                    check_flag = check_flag + 'True'
                else:
                    self.right_sm.setItem(right_index, 1, QtGui.QStandardItem(str("")))
                    check_flag = check_flag + 'False'
            if check_flag.__contains__('False'):
                QMessageBox.information(self, "Warning", "Some names are missing or invalid, "
                                                         "please enter a valid value. "
                                                         "And please do not use '_' as the first character"
                                                         " as it is used to define internal variables.",
                                        QMessageBox.Ok)
            return check_flag.__contains__('False')

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def check_dup_records(self):
        try:
            name_list = []
            for right_index in range(self.right_sm.rowCount()):
                name_text = self.right_sm.item(right_index, 1).text().strip()
                name_list.append(name_text)
            name_set = set(name_list)
            if len(name_list) != len(name_set):
                QMessageBox.information(self, "Warning", "Some names are duplicated.", QMessageBox.Ok)
                return False
            if len(name_list) == 0:
                return False
            return True

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def right_custom_right_menu(self):
        try:
            if self.right_sm.rowCount() != 0:
                indexes = self.righttableView.selectionModel().selectedIndexes()
                if len(indexes) != 0:
                    if self.right_sm.item(indexes[0].row(), 0).row() >= 0:
                        menu = QMenu(self)
                        opt1 = menu.addAction("Delete")
                        action = menu.exec_(QCursor.pos())
                        if action == opt1:
                            self.right_sm.removeRow(self.righttableView.currentIndex().row())

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def custom_right_menu(self):
        try:
            if self.left_sm.rowCount() != 0:
                indexes = self.lefttableView.selectionModel().selectedIndexes()
                selected = self.left_filter.mapToSource(indexes[0]).row()
                if len(indexes) != 0:
                    single_tag_selected = self.left_sm.item(selected, self.list_header_order['Tag']).text().strip()
                    tag_3_selected = self.left_sm.item(selected, self.list_header_order['Parent_3_Tag']).text().strip()
                    frame_selected = self.left_sm.item(selected, self.list_header_order['frame']).text().strip()
                    search_type_selected = self.left_sm.item(selected, self.list_header_order['search_type']).\
                        text().strip()
                    search_rule = search_type_selected
                    if '_3' in search_type_selected:
                        search_rule = search_type_selected.replace('_3', '')
                    value_selected = self.left_sm.item(selected, self.list_header_order[search_rule]).text()
                    menu = QMenu(self)
                    opt1 = menu.addAction("Select by {0}".format(search_rule))
                    action = menu.exec_(QCursor.pos())
                    if action == opt1:
                        self.select_by_attr_or_text(single_tag_selected, tag_3_selected, value_selected, frame_selected,
                                                    search_type_selected, search_rule)
        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    @staticmethod
    def get_expect_attribute(tag, parent_3_tags, attrs, search_type, search_rule):
        if search_type == 'Tag':
            return "(By.TAG_NAME, '{0}')".format(tag)
        if search_type == 'text':
            attrs_temp = attrs.replace("'", "\\'").replace('\n', '\\n')
            search_string = "(By.XPATH, '//{0}[text()=\"{1}\"]')".format(tag, attrs_temp)
            if '"' in attrs:
                attrs_temp = attrs.replace('"', '\\"').replace('\n', '\\n')
                search_string = '(By.XPATH, "//{0}[text()=\'{1}\']")'.format(tag, attrs_temp)
            return search_string
        if search_type == 'text_3':
            attrs_temp = attrs.replace("'", "\\'").replace('\n', '\\n')
            search_string = "(By.XPATH, '//{0}[text()=\"{1}\"]')".format(parent_3_tags, attrs_temp)
            if '"' in attrs:
                attrs_temp = attrs.replace('"', '\\"').replace('\n', '\\n')
                search_string = '(By.XPATH, "//{0}[text()=\'{1}\']")'.format(parent_3_tags, attrs_temp)
            return search_string
        if search_rule in CONFIG_DATA["preferred_attributes"] + ['main_xpath', 'full_xpath']:
            return "(By.XPATH, '//{0}{1}')".format(tag, attrs)
        if search_rule in CONFIG_DATA["preferred_attributes"] + ['main_xpath', 'full_xpath'] and '_3' in search_type:
            return "(By.XPATH, '//{0}{1}')".format(parent_3_tags, attrs)

    def select_by_attr_or_text(self, tag, parent_3_tags, attrs, frame, search_type, search_rule):
        try:
            rowcount = self.right_sm.rowCount()
            right_item = self.get_expect_attribute(tag, parent_3_tags, attrs, search_type, search_rule)
            self.right_sm.setItem(rowcount, 0, QtGui.QStandardItem(str(right_item)))
            self.right_sm.setItem(rowcount, 1, QtGui.QStandardItem(str("")))
            self.right_sm.setItem(rowcount, 2, QtGui.QStandardItem(frame))

        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def scan_browsers_note(self):
        print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        self.notelabel.setText("<font color=%s>%s</font>" % ('#FF0000', "Scanning, please wait a while ..."))

    def scan_browsers(self):
        time_delay = 0
        if self.time_delay.text() != '':
            time_delay = int(self.time_delay.text().strip())
        time.sleep(time_delay)
        t = threading.Thread(target=MyWindow.scan_browsers_contents, args=(self,))
        t.start()
        t.join()

    def scan_browsers_contents(self):
        try:
            if self.ObjectProperty.toPlainText().strip() == '':
                self.full_scan_flag = True
                runtime, url, data_path = self.scanner.html_info_by_driver(only_screen_shot_flag=False)
            else:
                self.full_scan_flag = False
                runtime, url, data_path = self.scanner.object_info_by_driver(self.ObjectProperty.toPlainText())
            loop = 1
            while not os.path.isfile(data_path + '\\HTMLObjects.xlsx') and loop < 10:
                time.sleep(1)
                loop = loop + 1

            self.notelabel.setText("<font color=%s>%s</font>" % ('#00CD00', "Scanning was completed"))

            self.urllistView_sm.insertRow(0, [QtGui.QStandardItem(str(runtime)), QtGui.QStandardItem(str(url))])

            print(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        except Exception as error:
            print(error.__str__())
            self.notelabel.setText(error.__str__())

    def load_url_data(self, url_file):
        try:
            self.urllistView_sm.clear()
            self.urllistView_sm.setHorizontalHeaderLabels(['Date', 'URL'])
            self.urllistView.setModel(self.urllistView_sm)
            self.urllistView.setColumnWidth(0, 200)
            self.urllistView.setColumnWidth(1, 730)

            if os.path.isfile(url_file):
                with open(url_file, "r") as f_output:
                    self.all_loaded_urls = f_output.readlines()
                for url_index in range(len(self.all_loaded_urls)):
                    url_line = self.all_loaded_urls[url_index].strip().replace('\n', '').replace('\r', '').strip()
                    self.urllistView_sm.setItem(url_index, 0, QtGui.QStandardItem(str(url_line.split(',')[0])))
                    self.urllistView_sm.setItem(url_index, 1, QtGui.QStandardItem(str(url_line.split(',')[1])))

        except Exception as error:
            print(error.__str__())

    def url_doubleclick(self):
        try:
            indexes = self.urllistView.selectionModel().selectedIndexes()
            selected = self.urllistView_sm.item(indexes[0].row(), 0).row()
            self.selected_url_date = self.urllistView_sm.item(selected, 0).text(). \
                replace(':', '_').replace('-', '_').replace(' ', '_').strip()
            self.selected_url = self.urllistView_sm.item(selected, 1).text().strip()
            name = parse_url(self.selected_url.strip())
            data_path = os.path.join(EXCEL_FOLDER_PATH, self.selected_url_date, name.strip())
            file_name = '{0}\\{1}'.format(data_path, 'HTMLObjects.xlsx')
            if os.path.isfile(file_name):
                url_data = pd.read_excel(file_name)
                url_data.fillna('', inplace=True)

                self.load_selected_data(data_path, url_data)
            else:
                QMessageBox.information(self, "Information", 'Did not find the expected scan result')
        except Exception as error:
            QMessageBox.information(self, "Information", self.tr(error.__str__()))

    def load_selected_data(self, data_path, url_data):
        self.left_filter.setSourceModel(self.left_sm)
        self.left_sm.clear()
        self.tabWidget.setCurrentIndex(0)
        self.class_name = ''

        left_title = ['Index', 'Tag', 'Attributes', 'Text', 'Frame']
        self.left_sm.setHorizontalHeaderLabels(left_title)
        self.lefttableView.setModel(self.left_filter)
        self.lefttableView.setColumnWidth(0, 50)
        self.lefttableView.setColumnWidth(1, 100)
        self.lefttableView.setColumnWidth(2, 500)
        self.lefttableView.setColumnWidth(3, 150)
        self.lefttableView.setColumnWidth(4, 120)

        right_title = ['Locator', 'Expected Name', 'Frame']
        self.right_sm.setHorizontalHeaderLabels(right_title)
        self.righttableView.setModel(self.right_sm)
        self.righttableView.setColumnWidth(0, 530)
        self.righttableView.setColumnWidth(1, 130)
        self.righttableView.setColumnWidth(2, 120)

        self.list_header_order = {}
        columns_list = ['Index', 'Tag_Mark', 'Attributes', 'text', 'frame'] + CONFIG_DATA["preferred_attributes"] \
                       + ['main_xpath', 'full_xpath', 'image', 'Tag', 'Parent_3_Tag', 'search_type', 'frame_index']
        for col_index in range(len(columns_list)):
            self.list_header_order[columns_list[col_index]] = col_index

        display_data = url_data[columns_list].copy()
        self.frame_data = url_data[url_data["Tag"] == 'iframe'].copy()
        self.showing_data = display_data[url_data["image"] != ''].copy()

        if self.full_scan_flag is False:
            self.showing_data["frame"] = self.showing_data["frame_index"].astype('str')
        else:
            self.showing_data["frame"] = self.showing_data["frame"].astype('str')

        df_array = self.showing_data.values
        for row in range(self.showing_data.shape[0]):
            for col in range(self.showing_data.shape[1]):
                self.left_sm.setItem(row, col, QtGui.QStandardItem(str(df_array[row, col])))
                if df_array[row, 13].strip() == 'N':
                    self.left_sm.item(row, col).setBackground(QBrush(QColor("yellow")))
        for col_index in range(len(columns_list)):
            if col_index > 4:
                self.lefttableView.setColumnHidden(col_index, True)

        self.tagcomboBox.clear()
        self.tag_list = list(set(self.showing_data["Tag_Mark"]))
        self.tag_list.sort()
        self.tagcomboBox.addItem("ALL")
        self.tagcomboBox.addItems(self.tag_list)
        if len(self.tag_list) > 0:
            self.tagcomboBox.setCurrentIndex(0)

        self.full_screen_name = os.path.join(data_path, "full_screen_shot.png")
        self.image_container.setText("")
        if os.path.isfile(self.full_screen_name):
            self.image_container.setPixmap(QPixmap(self.full_screen_name))
            self.image_container.setScaledContents(True)

    def filter_objects(self):
        search_option = self.searchcomboBox.currentText()
        if search_option == 'Tag':
            self.left_filter.setFilterKeyColumn(1)

        elif search_option == 'Attributes':
            self.left_filter.setFilterKeyColumn(2)
        else:
            self.left_filter.setFilterKeyColumn(3)

        search_text = self.searchText.text()
        self.left_filter.setFilterFixedString(search_text)
        self.left_filter.setFilterCaseSensitivity(False)

    def get_xy_coordinates(self):
        try:
            generated_excel_path = self.scanner.html_info_by_driver(only_screen_shot_flag=True)
            self.full_screen_name = os.path.join(generated_excel_path, "full_screen_shot.png")
            self.coordinates_image = cv2.imread(self.full_screen_name)
            cv2.namedWindow("image", 1)
            cv2.setMouseCallback("image", self.on_event_left_button_down)
            cv2.imshow("image", self.coordinates_image)
        except Exception:
            pass

    def on_event_left_button_down(self, event, x, y, flags, param):
        if event == cv2.EVENT_LBUTTONDOWN:
            if len(self.coordinates) < 2:
                self.coordinates.append([x, y])
                self.coordinates_image = cv2.imread(self.full_screen_name)
                for loc in range(len(self.coordinates)):
                    xy = "x: %d, y: %d" % (self.coordinates[loc][0], self.coordinates[loc][1])
                    cv2.putText(self.coordinates_image, xy, (self.coordinates[loc][0], self.coordinates[loc][1]),
                                cv2.FONT_HERSHEY_PLAIN, 1.0, (0, 0, 255), thickness=1)
                    cv2.imshow("image", self.coordinates_image)

            if len(self.coordinates) == 2:
                for loc in range(len(self.coordinates)):
                    xy = "x: %d, y: %d" % (self.coordinates[loc][0], self.coordinates[loc][1])
                    cv2.putText(self.coordinates_image, xy, (self.coordinates[loc][0], self.coordinates[loc][1]),
                                cv2.FONT_HERSHEY_PLAIN, 1.0, (0, 0, 255), thickness=1)

                cv2.rectangle(self.coordinates_image, (self.coordinates[0][0], self.coordinates[0][1]),
                              (self.coordinates[1][0], self.coordinates[1][1]), (0, 255, 0), 1)
                diff = "diff: (x: %d, y: %d)" % (int(self.coordinates[1][0] - self.coordinates[0][0]),
                                         int(self.coordinates[1][1] - self.coordinates[0][1]))
                cv2.putText(self.coordinates_image, diff, (int((self.coordinates[0][0] + self.coordinates[1][0])/2),
                                                           int((self.coordinates[0][1] + self.coordinates[1][1])/2)),
                            cv2.FONT_HERSHEY_PLAIN, 1.0, (0, 0, 255), thickness=1)
                cv2.imshow("image", self.coordinates_image)
                self.coordinates.clear()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyWindow()
    myWin.show()
    sys.exit(app.exec_())
