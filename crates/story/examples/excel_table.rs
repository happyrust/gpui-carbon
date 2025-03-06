use calamine::{open_workbook, Reader, Xlsx};
use fake::faker::name;
use gpui::{
    div, hsla, impl_actions, px, App, AppContext, Application, BorrowAppContext, Context, Edges,
    Entity, EventEmitter, FocusHandle, Focusable, InteractiveElement, IntoElement, ParentElement,
    Pixels, Render, SharedString, Size, Styled, Subscription, Window,
};
use gpui_component::{
    button::Button,
    dock::{DockArea, DockItem, DockPlacement, Panel, PanelEvent, PanelView},
    dropdown::{Dropdown, DropdownEvent},
    h_flex,
    input::TextInput,
    label::Label,
    notification::{Notification, NotificationType},
    popup_menu::PopupMenuExt,
    table::{self, Table, TableDelegate, TableEvent},
    v_flex, ContextModal, Sizable,
};
use rusqlite::{params, Connection, Result as SqliteResult};
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::ops::Range;
use std::path::PathBuf;
use std::sync::Arc;
use story::{Assets, Story};

#[derive(Clone, Debug)]
struct ExcelRow {
    id: usize,
    data: HashMap<String, String>,
}

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct SetProjectType(String);

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct SetRoadType(String);

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct ChangeSheet {
    sheet_name: String,
}

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct UpdateResultTable;

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct UpdateTypesUIEvent;

impl_actions!(
    excel_table,
    [
        SetProjectType,
        SetRoadType,
        ChangeSheet,
        UpdateResultTable,
        UpdateTypesUIEvent
    ]
);

struct ExcelTableDelegate {
    rows: Vec<ExcelRow>,
    columns: Vec<String>,
    size: Size<Pixels>,
    loop_selection: bool,
    col_resize: bool,
    col_order: bool,
    col_sort: bool,
    col_selection: bool,
    loading: bool,
    full_loading: bool,
    fixed_cols: bool,
    eof: bool,
    visible_rows: Range<usize>,
    visible_cols: Range<usize>,
}

impl ExcelTableDelegate {
    fn new() -> Self {
        Self {
            rows: Vec::new(),
            columns: Vec::new(),
            size: Size::default(),
            loop_selection: true,
            col_resize: true,
            col_order: true,
            col_sort: true,
            col_selection: true,
            loading: false,
            full_loading: false,
            fixed_cols: false,
            eof: false,
            visible_rows: Range::default(),
            visible_cols: Range::default(),
        }
    }

    fn set_data(&mut self, columns: Vec<String>, data: Vec<HashMap<String, String>>) {
        self.columns = columns;
        self.rows = data
            .into_iter()
            .enumerate()
            .map(|(id, data)| ExcelRow { id, data })
            .collect();
        self.eof = true;
        self.loading = false;
        self.full_loading = false;
    }
}

impl TableDelegate for ExcelTableDelegate {
    fn cols_count(&self, _: &App) -> usize {
        self.columns.len()
    }

    fn rows_count(&self, _: &App) -> usize {
        self.rows.len()
    }

    fn col_name(&self, col_ix: usize, _: &App) -> SharedString {
        if let Some(col) = self.columns.get(col_ix) {
            col.clone().into()
        } else {
            "--".into()
        }
    }

    fn col_width(&self, _: usize, _: &App) -> Pixels {
        120.0.into()
    }

    fn col_padding(&self, _: usize, _: &App) -> Option<Edges<Pixels>> {
        Some(Edges::all(px(4.)))
    }

    fn col_fixed(&self, col_ix: usize, _: &App) -> Option<table::ColFixed> {
        if !self.fixed_cols {
            return None;
        }

        if col_ix < 2 {
            Some(table::ColFixed::Left)
        } else {
            None
        }
    }

    fn can_resize_col(&self, _: usize, _: &App) -> bool {
        self.col_resize
    }

    fn can_select_col(&self, _: usize, _: &App) -> bool {
        self.col_selection
    }

    fn render_th(
        &self,
        col_ix: usize,
        _: &mut Window,
        cx: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        div().child(self.col_name(col_ix, cx))
    }

    fn render_td(
        &self,
        row_ix: usize,
        col_ix: usize,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        let row = &self.rows[row_ix];
        let col_name = &self.columns[col_ix];

        div().child(row.data.get(col_name).cloned().unwrap_or_default())
    }

    fn can_loop_select(&self, _: &App) -> bool {
        self.loop_selection
    }

    fn can_move_col(&self, _: usize, _: &App) -> bool {
        self.col_order
    }

    fn move_col(
        &mut self,
        col_ix: usize,
        to_ix: usize,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) {
        let col = self.columns.remove(col_ix);
        self.columns.insert(to_ix, col);
    }

    fn loading(&self, _: &App) -> bool {
        self.full_loading
    }

    fn can_load_more(&self, _: &App) -> bool {
        !self.loading && !self.eof
    }

    fn visible_rows_changed(
        &mut self,
        visible_range: Range<usize>,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) {
        self.visible_rows = visible_range;
    }

    fn visible_cols_changed(
        &mut self,
        visible_range: Range<usize>,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) {
        self.visible_cols = visible_range;
    }
}

#[derive(Clone)]
pub struct ExcelStory {
    dock_area: Entity<DockArea>,
    table: Entity<Table<ExcelTableDelegate>>,
    current_sheet: Option<String>,
    required_columns: Vec<String>,
    focus_handle: FocusHandle,
    project_type: String,
    road_type: String,
    db_path: String,
}

impl Story for ExcelStory {
    fn title() -> &'static str {
        "Excel Table"
    }

    fn description() -> &'static str {
        "A table component that can display Excel file data with configurable columns"
    }

    fn klass() -> &'static str {
        "ExcelStory"
    }

    fn new_view(window: &mut Window, cx: &mut App) -> Entity<Self> {
        cx.new(|cx| Self::new(window, cx))
    }
}

impl Focusable for ExcelStory {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl EventEmitter<UpdateResultTable> for ExcelStory {}
impl EventEmitter<UpdateTypesUIEvent> for ExcelStory {}

impl ExcelStory {
    pub fn view(window: &mut Window, cx: &mut Context<Self>) -> Entity<Self> {
        cx.new(|cx| Self::new(window, cx))
    }

    fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let focus_handle = cx.focus_handle();

        let current_sheet = None;

        let required_columns = vec![
            "序号".to_string(),
            "编码".to_string(),
            "名称及规格".to_string(),
            "单位".to_string(),
            "数量".to_string(),
            "市场价".to_string(),
            "合计".to_string(),
        ];

        let delegate = ExcelTableDelegate::new();
        let table = cx.new(|cx| Table::new(delegate, window, cx));

        // Initialize dock area
        let dock_area = cx.new(|cx| DockArea::new("excel-table", Some(1), window, cx));

        // Create data directory if it doesn't exist
        let data_dir = "data";
        fs::create_dir_all(data_dir).expect("Failed to create data directory");
        let db_path = format!("{}/excel_data.db", data_dir);

        // Initialize database
        Self::init_carbonref_database(&db_path).expect("Failed to initialize 人材机数据库 ");

        let story_entity = cx.entity();

        let weak_dock_area = dock_area.downgrade();
        // Configure dock area layout
        dock_area.update(cx, |dock_area, cx| {
            // Left panel (Form panel)
            let left_panel = DockItem::tab(
                ParamFormPane::view(story_entity.clone(), window, cx),
                &weak_dock_area,
                window,
                cx,
            );

            // Center panel (Result panel)
            let center_panel = DockItem::tab(
                CarbonResultPanel::view(story_entity.clone(), window, cx),
                &weak_dock_area,
                window,
                cx,
            );

            // Right panel (Table panel)
            let right_panel = DockItem::tab(
                InputeTablePanel::view(story_entity.clone(), window, cx),
                &weak_dock_area,
                window,
                cx,
            );

            // Set panels in dock area
            dock_area.set_center(center_panel, window, cx);
            dock_area.set_left_dock(left_panel, Some(px(300.)), true, window, cx);
            dock_area.set_right_dock(right_panel, Some(px(600.)), true, window, cx);
        });

        let excel_story = Self {
            dock_area,
            table,
            current_sheet,
            required_columns,
            focus_handle,
            project_type: String::new(),
            road_type: String::new(),
            db_path,
        };

        // Load resource data
        excel_story.load_resource_data(window, cx);

        excel_story
    }

    fn init_carbonref_database(db_path: &str) -> SqliteResult<()> {
        let conn = Connection::open(db_path)?;

        // Create sheets table
        conn.execute(
            "CREATE TABLE IF NOT EXISTS sheets (
                id INTEGER PRIMARY KEY,
                name TEXT NOT NULL UNIQUE
            )",
            [],
        )?;

        // Create data table
        conn.execute(
            "CREATE TABLE IF NOT EXISTS excel_data (
                id INTEGER PRIMARY KEY,
                sheet_id INTEGER,
                序号 TEXT,
                编码 TEXT,
                名称及规格 TEXT,
                单位 TEXT,
                数量 TEXT,
                市场价 TEXT,
                合计 TEXT,
                FOREIGN KEY(sheet_id) REFERENCES sheets(id)
            )",
            [],
        )?;

        // Create resource tables
        conn.execute(
            "CREATE TABLE IF NOT EXISTS labor (
                code TEXT PRIMARY KEY,
                name TEXT,
                specification TEXT,
                unit TEXT,
                carbon_factor REAL
            )",
            [],
        )?;

        conn.execute(
            "CREATE TABLE IF NOT EXISTS material (
                code TEXT PRIMARY KEY,
                name TEXT,
                specification TEXT,
                unit TEXT,
                carbon_factor REAL
            )",
            [],
        )?;

        conn.execute(
            "CREATE TABLE IF NOT EXISTS machine (
                code TEXT PRIMARY KEY,
                name TEXT,
                specification TEXT,
                unit TEXT,
                carbon_factor REAL
            )",
            [],
        )?;

        Ok(())
    }

    fn change_sheet(&mut self, sheet_name: String, window: &mut Window, cx: &mut Context<Self>) {
        match self.load_sheet_data(&sheet_name) {
            Ok((headers, data)) => {
                let data_len = data.len();
                self.table.update(cx, |table, cx| {
                    table.delegate_mut().set_data(headers, data);
                    table.refresh(cx);
                });

                window.push_notification(
                    Notification::new(format!(
                        "已切换到工作表 '{}', 共 {} 行数据",
                        sheet_name, data_len
                    ))
                    .with_type(NotificationType::Success),
                    cx,
                );
            }
            Err(e) => {
                window.push_notification(
                    Notification::new(format!("加载工作表数据失败: {}", e))
                        .with_type(NotificationType::Error),
                    cx,
                );
            }
        }
    }

    /// 加载指定工作表的数据
    /// 参数:
    ///   - sheet_name: 工作表名称
    /// 返回:
    ///   - Ok((headers, data)): headers为表头列名, data为表格数据
    ///   - Err: 数据库操作错误
    fn load_sheet_data(
        &self,
        sheet_name: &str,
    ) -> SqliteResult<(Vec<String>, Vec<HashMap<String, String>>)> {
        let conn = Connection::open(&self.db_path)?;

        // Get sheet_id
        let sheet_id: i64 = conn.query_row(
            "SELECT id FROM sheets WHERE name = ?",
            params![sheet_name],
            |row| row.get(0),
        )?;

        // Get data for the sheet
        let mut stmt = conn.prepare(
            "SELECT 序号, 编码, 名称及规格, 单位, 数量, 市场价, 合计 
             FROM excel_data 
             WHERE sheet_id = ?
             ORDER BY id",
        )?;

        let mut data = Vec::new();
        let mut rows = stmt.query(params![sheet_id])?;

        while let Some(row) = rows.next()? {
            let mut row_data = HashMap::new();
            for (i, col) in self.required_columns.iter().enumerate() {
                let value: String = row.get(i)?;
                row_data.insert(col.clone(), value);
            }
            data.push(row_data);
        }

        Ok((self.required_columns.clone(), data))
    }

    /// 从数据库中提取所有工程类型和道路类型
    fn extract_project_and_road_types(&self) -> SqliteResult<(Vec<String>, Vec<String>)> {
        let conn = Connection::open(&self.db_path)?;
        let mut stmt = conn.prepare("SELECT name FROM sheets ORDER BY id")?;

        let mut project_types = std::collections::HashSet::new();
        let mut road_types = std::collections::HashSet::new();

        let rows = stmt.query_map([], |row| row.get::<_, String>(0))?;

        for sheet_name in rows.flatten() {
            if let Some((project_type, road_type)) = Self::parse_sheet_name(&sheet_name) {
                project_types.insert(project_type);
                road_types.insert(road_type);
            }
        }

        let mut project_types: Vec<_> = project_types.into_iter().collect();
        let mut road_types: Vec<_> = road_types.into_iter().collect();
        project_types.sort();
        road_types.sort();

        Ok((project_types, road_types))
    }

    /// 从工作表名称中解析出工程类型和道路类型
    fn parse_sheet_name(sheet_name: &str) -> Option<(String, String)> {
        // 假设格式为: "xxx【project_type road_type】"
        // todo 空格划分？
        let mut names = sheet_name.trim().split_whitespace().into_iter();
        let project_type = names.next()?.to_string();
        let road_type = names.fold(None, |mut acc: Option<String>, s| {
            if acc.is_none() {
                acc = Some(s.to_string());
            } else {
                acc.as_mut().unwrap().push_str(s);
            }
            acc
        })?;

        // dbg!()

        Some((project_type, road_type))
    }

    fn update_types_ui(&self, window: &mut Window, cx: &mut Context<Self>) {
        if let Ok((project_types, road_types)) = self.extract_project_and_road_types() {
            cx.emit(UpdateTypesUIEvent);
            //send event to ParamFormPane
            // 更新工程类型下拉框
            // if let Some(panel) = self.dock_area.find_("ParamFormPane", cx) {
            //     panel.update(cx, |panel, cx| {
            //         panel.project_type_dropdown.update(cx, |dropdown, cx| {
            //             dropdown.set_items(
            //                 project_types.into_iter().map(SharedString::from).collect(),
            //                 window,
            //                 cx,
            //             );
            //         });
            //         panel.road_type_dropdown.update(cx, |dropdown, cx| {
            //             dropdown.set_items(
            //                 road_types.into_iter().map(SharedString::from).collect(),
            //                 window,
            //                 cx,
            //             );
            //         });
            //     });
            // }
        }
    }

    fn load_excel(&mut self, file_path: &str, window: &mut Window, cx: &mut Context<Self>) {
        let path = PathBuf::from(file_path);

        match open_workbook::<Xlsx<_>, _>(&path) {
            Ok(mut workbook) => {
                let sheet_names = workbook.sheet_names().to_vec();

                // Open database connection
                match Connection::open(&self.db_path) {
                    Ok(mut conn) => {
                        // Start transaction
                        let tx = conn.transaction().unwrap();

                        // Don't clear all data, we'll handle updates per sheet
                        let mut success = false;

                        for sheet_name in &sheet_names {
                            if let Ok(range) = workbook.worksheet_range(sheet_name) {
                                if !range.is_empty() {
                                    let rows: Vec<_> = range.rows().collect();

                                    // Find header row
                                    let mut header_columns = HashMap::new();
                                    if let Some(header_idx) = rows.iter().position(|row| {
                                        header_columns.clear();
                                        for (col_idx, cell) in row.iter().enumerate() {
                                            let cell_value = cell.to_string().trim().to_string();
                                            if !cell_value.is_empty() {
                                                header_columns.insert(col_idx, cell_value);
                                            }
                                        }

                                        self.required_columns.iter().all(|required_col| {
                                            header_columns.values().any(|col| col == required_col)
                                        })
                                    }) {
                                        // Check if sheet exists and get its ID
                                        let sheet_id: Option<i64> = tx
                                            .query_row(
                                                "SELECT id FROM sheets WHERE name = ?",
                                                params![sheet_name],
                                                |row| row.get(0),
                                            )
                                            .ok();

                                        // If sheet exists, delete its data
                                        // if let Some(id) = sheet_id {
                                        //     tx.execute(
                                        //         "DELETE FROM excel_data WHERE sheet_id = ?",
                                        //         params![id],
                                        //     )
                                        //     .unwrap();
                                        //     tx.execute(
                                        //         "DELETE FROM sheets WHERE id = ?",
                                        //         params![id],
                                        //     )
                                        //     .unwrap();
                                        // }

                                        // Insert or update sheet
                                        tx.execute(
                                            "INSERT INTO sheets (name) VALUES (?)",
                                            params![sheet_name],
                                        );
                                        let sheet_id = tx.last_insert_rowid();

                                        // Create column mapping
                                        let mut column_mapping = HashMap::new();
                                        for (col_idx, header) in header_columns.iter() {
                                            if self.required_columns.contains(header) {
                                                column_mapping.insert(*col_idx, header.clone());
                                            }
                                        }

                                        // Insert data
                                        for row in rows.iter().skip(header_idx + 1) {
                                            let mut row_data = HashMap::new();
                                            let mut valid_columns = 0;

                                            for required_col in &self.required_columns {
                                                if let Some(col_idx) = column_mapping
                                                    .iter()
                                                    .find(|(_, header)| *header == required_col)
                                                    .map(|(&idx, _)| idx)
                                                {
                                                    if let Some(cell) = row.get(col_idx) {
                                                        let cell_value =
                                                            cell.to_string().trim().to_string();
                                                        if !cell_value.is_empty() {
                                                            row_data.insert(
                                                                required_col.clone(),
                                                                cell_value,
                                                            );
                                                            valid_columns += 1;
                                                        }
                                                    }
                                                }
                                            }

                                            if valid_columns > 0 {
                                                tx.execute(
                                                    "INSERT INTO excel_data (
                                                        sheet_id, 序号, 编码, 名称及规格, 单位, 数量, 市场价, 合计
                                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                                    params![
                                                        sheet_id,
                                                        row_data.get("序号").unwrap_or(&String::new()),
                                                        row_data.get("编码").unwrap_or(&String::new()),
                                                        row_data.get("名称及规格").unwrap_or(&String::new()),
                                                        row_data.get("单位").unwrap_or(&String::new()),
                                                        row_data.get("数量").unwrap_or(&String::new()),
                                                        row_data.get("市场价").unwrap_or(&String::new()),
                                                        row_data.get("合计").unwrap_or(&String::new()),
                                                    ],
                                                ).unwrap();
                                            }
                                        }
                                        success = true;
                                    }
                                }
                            }
                        }

                        if success {
                            // Commit transaction
                            tx.commit().unwrap();

                            // Get available sheets
                            let mut stmt =
                                conn.prepare("SELECT name FROM sheets ORDER BY id").unwrap();
                            let sheet_names: Vec<String> = stmt
                                .query_map([], |row| row.get(0))
                                .unwrap()
                                .map(|r| r.unwrap())
                                .collect();

                            // self.project_type
                            // Update dropdowns with new project and road types
                            self.update_types_ui(window, cx);

                            // Load first sheet
                            if let Some(first_sheet) = sheet_names.first() {
                                self.current_sheet = Some(first_sheet.clone());
                                self.change_sheet(first_sheet.clone(), window, cx);
                            }

                            window.push_notification(
                                Notification::new("Excel 数据已成功导入到数据库")
                                    .with_type(NotificationType::Success),
                                cx,
                            );
                        } else {
                            window.push_notification(
                                Notification::new("未找到任何有效的工作表数据")
                                    .with_type(NotificationType::Error),
                                cx,
                            );
                        }
                    }
                    Err(e) => {
                        window.push_notification(
                            Notification::new(format!("数据库连接失败: {}", e))
                                .with_type(NotificationType::Error),
                            cx,
                        );
                    }
                }
            }
            Err(e) => {
                window.push_notification(
                    Notification::new(format!("打开 Excel 文件失败: {}", e))
                        .with_type(NotificationType::Error),
                    cx,
                );
            }
        }
    }

    fn load_resource_data(&self, window: &mut Window, cx: &mut Context<Self>) {
        let file_path = PathBuf::from("assets/excel/人材机数据库.xlsx");

        match open_workbook::<Xlsx<_>, _>(&file_path) {
            Ok(mut workbook) => {
                // Open database connection
                match Connection::open(&self.db_path) {
                    Ok(mut conn) => {
                        // Start transaction
                        let tx = conn.transaction().unwrap();

                        // Clear existing data
                        tx.execute("DELETE FROM labor", []).unwrap();
                        tx.execute("DELETE FROM material", []).unwrap();
                        tx.execute("DELETE FROM machine", []).unwrap();

                        let required_columns =
                            vec!["编码", "名称", "规格型号", "单位", "单位碳排放因子"];

                        let sheet_table_mapping = vec![
                            ("人工数据", "labor"),
                            ("材料数据", "material"),
                            ("机械数据", "machine"),
                        ];

                        let mut success = false;

                        for (sheet_name, table_name) in sheet_table_mapping {
                            // dbg!(sheet_name, table_name);
                            if let Ok(range) = workbook.worksheet_range(sheet_name) {
                                if !range.is_empty() {
                                    let rows: Vec<_> = range.rows().collect();

                                    // Find header row and column indices
                                    let mut column_indices = HashMap::new();
                                    if let Some(header_idx) = rows.iter().position(|row| {
                                        column_indices.clear();
                                        for (col_idx, cell) in row.iter().enumerate() {
                                            let cell_value = cell.to_string().trim().to_string();
                                            if required_columns.contains(&cell_value.as_str()) {
                                                column_indices.insert(cell_value, col_idx);
                                            }
                                        }
                                        required_columns
                                            .iter()
                                            .all(|col| column_indices.contains_key(*col))
                                    }) {
                                        // dbg!(&column_indices);
                                        // Insert data
                                        let insert_sql = format!(
                                            "INSERT INTO {} (code, name, specification, unit, carbon_factor) 
                                             VALUES (?, ?, ?, ?, ?)",
                                            table_name
                                        );

                                        let mut stmt = tx.prepare(&insert_sql).unwrap();

                                        for row in rows.iter().skip(header_idx + 1) {
                                            let code = row
                                                .get(column_indices["编码"])
                                                .map(|c| c.to_string())
                                                .unwrap_or_default();

                                            if !code.is_empty() {
                                                let name = row
                                                    .get(column_indices["名称"])
                                                    .map(|c| c.to_string())
                                                    .unwrap_or_default();

                                                let specification = row
                                                    .get(column_indices["规格型号"])
                                                    .map(|c| c.to_string())
                                                    .unwrap_or_default();

                                                let unit = row
                                                    .get(column_indices["单位"])
                                                    .map(|c| c.to_string())
                                                    .unwrap_or_default();

                                                // if empty 1.0
                                                let carbon_factor = row
                                                    .get(column_indices["单位碳排放因子"])
                                                    .and_then(|c| c.to_string().parse::<f64>().ok())
                                                    .unwrap_or(1.0);

                                                stmt.execute(params![
                                                    code,
                                                    name,
                                                    specification,
                                                    unit,
                                                    carbon_factor,
                                                ])
                                                .unwrap();
                                            }
                                        }
                                        success = true;
                                    }
                                }
                            }
                        }

                        if success {
                            tx.commit().unwrap();
                            window.push_notification(
                                Notification::new("人材机数据库已成功导入")
                                    .with_type(NotificationType::Success),
                                cx,
                            );
                        } else {
                            window.push_notification(
                                Notification::new("未找到有效的人材机数据")
                                    .with_type(NotificationType::Error),
                                cx,
                            );
                        }
                    }
                    Err(e) => {
                        window.push_notification(
                            Notification::new(format!("数据库连接失败: {}", e))
                                .with_type(NotificationType::Error),
                            cx,
                        );
                    }
                }
            }
            Err(e) => {
                window.push_notification(
                    Notification::new(format!("打开人材机数据库文件失败: {}", e))
                        .with_type(NotificationType::Error),
                    cx,
                );
            }
        }
    }

    fn on_change_sheet(
        &mut self,
        action: &ChangeSheet,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        self.change_sheet(action.sheet_name.clone(), window, cx);
        cx.emit(UpdateResultTable);
    }
}

impl Render for ExcelStory {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        v_flex()
            .size_full()
            // .on_action(cx.listener(Self::on_set_project_type))
            // .on_action(cx.listener(Self::on_set_road_type))
            .on_action(cx.listener(Self::on_change_sheet))
            .child(self.dock_area.clone())
    }
}

struct ParamFormPane {
    story: Entity<ExcelStory>,
    focus_handle: FocusHandle,
    project_type_dropdown: Entity<Dropdown<Vec<SharedString>>>,
    road_type_dropdown: Entity<Dropdown<Vec<SharedString>>>,
    file_path_input: Entity<TextInput>,
}

impl ParamFormPane {
    pub fn view(
        story: Entity<ExcelStory>,
        window: &mut Window,
        cx: &mut Context<DockArea>,
    ) -> Entity<Self> {
        let file_path_input = cx.new(|cx| {
            let mut input = TextInput::new(window, cx);
            input.set_text("assets/excel/指标汇总.xlsx", window, cx);
            input
        });

        let project_types = vec!["道路工程".into(), "交通工程".into()];
        let project_type_dropdown = cx.new(|cx| {
            Dropdown::new("project-type", project_types, None, window, cx)
                .small()
                .placeholder("请选择工程类型")
        });

        let road_types = vec!["主干路 62 m2".into(), "次干路59cm m2".into()];
        let road_type_dropdown = cx.new(|cx| {
            Dropdown::new("road-type", road_types, None, window, cx)
                .small()
                .placeholder("请选择道路类型")
        });

        let view = cx.new(|cx| {
            let story = story.clone();
            let project_type_dropdown = project_type_dropdown.clone();
            let road_type_dropdown = road_type_dropdown.clone();
            let file_path_input = file_path_input.clone();

            let panel = Self {
                story: story.clone(),
                focus_handle: cx.focus_handle(),
                project_type_dropdown: project_type_dropdown.clone(),
                road_type_dropdown: road_type_dropdown.clone(),
                file_path_input: file_path_input.clone(),
            };

            // Subscribe to dropdown events
            cx.subscribe_in(&project_type_dropdown, window, {
                let story = story.clone();
                move |_, _dropdown, event: &DropdownEvent<Vec<SharedString>>, _window, cx| {
                    if let DropdownEvent::Confirm(Some(value)) = event {
                        story.update(cx, |story, cx| {
                            story.project_type = value.to_string();
                            // 只有当两个类型都已选择时才发送更新事件
                            if !story.project_type.is_empty() && !story.road_type.is_empty() {
                                cx.emit(UpdateResultTable);
                            }
                        });
                    }
                }
            })
            .detach();

            cx.subscribe_in(&road_type_dropdown, window, {
                let story = story.clone();
                move |_, _dropdown, event: &DropdownEvent<Vec<SharedString>>, _window, cx| {
                    if let DropdownEvent::Confirm(Some(value)) = event {
                        story.update(cx, |story, cx| {
                            story.road_type = value.to_string();
                            // 只有当两个类型都已选择时才发送更新事件
                            if !story.project_type.is_empty() && !story.road_type.is_empty() {
                                cx.emit(UpdateResultTable);
                            }
                        });
                    }
                }
            })
            .detach();

            // UpdateTypesUIEvent
            cx.subscribe_in(&story, window, {
                let story = story.clone();
                move |_, _dropdown, event: &UpdateTypesUIEvent, _window, cx| {
                    dbg!("UpdateTypesUIEvent");
                    // if let DropdownEvent::Confirm(Some(value)) = event {
                    //     story.update(cx, |story, cx| {
                    //         story.road_type = value.to_string();
                    //         // 只有当两个类型都已选择时才发送更新事件
                    //         if !story.project_type.is_empty() && !story.road_type.is_empty() {
                    //             cx.emit(UpdateResultTable);
                    //         }
                    //     });
                    // }
                }
            })
            .detach();

            panel
        });

        view
    }
}

impl EventEmitter<PanelEvent> for ParamFormPane {}

impl Panel for ParamFormPane {
    fn panel_name(&self) -> &'static str {
        "ParamFormPane"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "配置".into_any_element()
    }
}

impl Focusable for ParamFormPane {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for ParamFormPane {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        // let this = self.story.read(cx);
        let file_path_input = self.file_path_input.clone();
        div()
            .flex()
            .flex_col()
            // .min_w(px(100.))
            .w_auto()
            .p_4()
            .bg(hsla(0.0, 0.0, 0.96, 1.0))
            .child(
                v_flex()
                    .gap_4()
                    .child(
                        v_flex()
                            .gap_2()
                            .child(Label::new("Excel 文件路径"))
                            .child(
                                div().w_full().child(
                                    h_flex()
                                        .gap_1()
                                        .child(div().flex_grow().child(file_path_input.clone()))
                                        .child(Button::new("select-file").label("...").on_click({
                                            let file_path_input = file_path_input.clone();
                                            move |_, window, cx| {
                                                let file_path_input = file_path_input.clone();
                                                window
                                                    .spawn(cx, |mut awc| async move {
                                                        if let Some(path) =
                                                            rfd::AsyncFileDialog::new()
                                                                .add_filter(
                                                                    "Excel files",
                                                                    &["xlsx"],
                                                                )
                                                                .set_title("选择 Excel 文件")
                                                                .pick_file()
                                                                .await
                                                        {
                                                            let path_str = path
                                                                .path()
                                                                .to_string_lossy()
                                                                .to_string();
                                                            // awc.update_entity(handle, update)
                                                            awc.update(|win, cx| {
                                                                file_path_input.update(
                                                                    cx,
                                                                    |input, cx| {
                                                                        input.set_text(
                                                                            &path_str, win, cx,
                                                                        );
                                                                    },
                                                                );
                                                            });
                                                        }
                                                    })
                                                    .detach();
                                            }
                                        })),
                                ),
                            )
                            .child(
                                h_flex()
                                    .gap_2()
                                    .child(
                                        Button::new("load")
                                            .label("导入Excel")
                                            .on_click({
                                                let story = self.story.clone();
                                                let file_path_input = self.file_path_input.clone();
                                                move |_, window, cx| {
                                                    let file_path = file_path_input.read(cx).text();
                                                    story.update(cx, |this, cx| {
                                                        this.load_excel(&file_path, window, cx);
                                                    });
                                                }
                                            })
                                            .flex_grow(),
                                    )
                                    .child({
                                        let this = self.story.read(cx);
                                        let current_sheet_label = this
                                            .current_sheet
                                            .as_deref()
                                            .unwrap_or("选择工作表")
                                            .to_string();
                                        let button = Button::new("sheet")
                                            .label(current_sheet_label)
                                            .flex_grow();

                                        let db_path = this.db_path.clone();
                                        button.popup_menu(move |menu, _, _| {
                                            let mut menu = menu;
                                            if let Ok(conn) = Connection::open(&db_path) {
                                                if let Ok(mut stmt) = conn
                                                    .prepare("SELECT name FROM sheets ORDER BY id")
                                                {
                                                    if let Ok(rows) = stmt.query_map([], |row| {
                                                        row.get::<_, String>(0)
                                                    }) {
                                                        for sheet_name in rows.flatten() {
                                                            menu = menu.menu(
                                                                sheet_name.clone(),
                                                                Box::new(ChangeSheet {
                                                                    sheet_name: sheet_name.clone(),
                                                                }),
                                                            );
                                                        }
                                                    }
                                                }
                                            }
                                            menu
                                        })
                                    }),
                            ),
                    )
                    .child(
                        v_flex()
                            .gap_2()
                            .child(Label::new("工程类型"))
                            .child(self.project_type_dropdown.clone()),
                    )
                    .child(
                        v_flex()
                            .gap_2()
                            .child(Label::new("道路类型"))
                            .child(self.road_type_dropdown.clone()),
                    ),
            )
    }
}

struct InputeTablePanel {
    story: Entity<ExcelStory>,
    focus_handle: FocusHandle,
}

impl InputeTablePanel {
    pub fn view(
        story: Entity<ExcelStory>,
        _window: &mut Window,
        cx: &mut Context<DockArea>,
    ) -> Entity<Self> {
        let view = cx.new(|cx| Self {
            story,
            focus_handle: cx.focus_handle(),
        });
        view
    }
}

impl EventEmitter<PanelEvent> for InputeTablePanel {}

impl Panel for InputeTablePanel {
    fn panel_name(&self) -> &'static str {
        "TablePanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "数据".into_any_element()
    }
}

impl Focusable for InputeTablePanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for InputeTablePanel {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let this = self.story.read(cx);
        div().size_full().child(this.table.clone())
    }
}

#[derive(Clone, Default, Debug)]
struct ResultRow {
    序号: String,
    编码: String,
    名称及规格: String,
    项目名称: String,
    单位: String,
    可研估算: String,
    碳排放指数: String,
    人工: String,
    材料: String,
    机械: String,
    小计: String,
}

#[derive(Clone)]
struct ResultTableDelegate {
    total_rows: Vec<ResultRow>,
    sub_rows: Vec<ResultRow>,
    columns: Vec<String>,
}

impl ResultTableDelegate {
    fn new() -> Self {
        let columns = vec![
            "序号".to_string(),
            "项目名称".to_string(),
            "单位".to_string(),
            "可研估算".to_string(),
            "碳排放指数".to_string(),
            "人工".to_string(),
            "材料".to_string(),
            "机械".to_string(),
            "小计".to_string(),
        ];

        Self {
            total_rows: vec![],
            sub_rows: vec![],
            columns,
        }
    }

    fn update_data(
        &mut self,
        project_type: &str,
        road_type: &str,
        db_path: &str,
    ) -> SqliteResult<()> {
        // 清空现有数据
        self.total_rows.clear();
        self.sub_rows.clear();

        // 如果工程类型或道路类型为空，直接返回
        if project_type.is_empty() || road_type.is_empty() {
            return Ok(());
        }

        let conn = Connection::open(db_path)?;
        let sheet_name = format!("{project_type} {road_type}");

        let sheet_id = if project_type == "道路工程" { 1 } else { 2 };

        // 查询所有数据，包括碳排放因子
        let mut stmt = conn.prepare(
            "
            SELECT DISTINCT 
                e.序号,
                e.编码,
                e.名称及规格,
                e.单位,
                e.数量,
                COALESCE(l.carbon_factor, 1) as labor_factor,
                COALESCE(m.carbon_factor, 1) as material_factor,
                COALESCE(mc.carbon_factor, 1) as machine_factor
             FROM excel_data e
             LEFT JOIN labor l ON e.编码 = l.code
             LEFT JOIN material m ON e.编码 = m.code
             LEFT JOIN machine mc ON e.编码 = mc.code
             WHERE e.sheet_id = :sheet_id
             AND e.编码 IS NOT NULL
             AND e.编码 != ''
             ORDER BY e.id;
            ",
        )?;

        let rows = stmt.query_map([sheet_id], |row| {
            let 序号: String = row.get(0)?;
            let 编码: String = row.get(1)?;
            let 名称及规格: String = row.get(2)?;
            let 单位: String = row.get(3)?;
            let 数量: String = row.get(4)?;
            let 数量_float = 数量.parse::<f64>().unwrap_or(0.0);

            let labor_factor: f64 = row.get(5)?;
            let material_factor: f64 = row.get(6)?;
            let machine_factor: f64 = row.get(7)?;

            // 计算碳排放量
            let labor_emission = labor_factor * 数量_float;
            let material_emission = material_factor * 数量_float;
            let machine_emission = machine_factor * 数量_float;
            let total_emission = labor_emission + material_emission + machine_emission;

            Ok(ResultRow {
                序号,
                编码: 编码.clone(),
                名称及规格: 名称及规格.clone(),
                项目名称: format!("{} {}", 编码, 名称及规格),
                单位,
                可研估算: 数量,
                碳排放指数: format!("{:.2}", total_emission / 数量_float),
                人工: format!("{:.4}", labor_emission),
                材料: format!("{:.4}", material_emission),
                机械: format!("{:.4}", machine_emission),
                小计: format!("{:.4}", total_emission),
            })
        })?;

        // 添加工程类型行
        self.total_rows.push(ResultRow {
            序号: "一".to_string(),
            编码: "".to_string(),
            名称及规格: "".to_string(),
            项目名称: project_type.to_string(),
            单位: "".to_string(),
            可研估算: "".to_string(),
            碳排放指数: "".to_string(),
            人工: "".to_string(),
            材料: "".to_string(),
            机械: "".to_string(),
            小计: "".to_string(),
        });

        // 添加道路类型行
        let mut total_labor = 0.0;
        let mut total_material = 0.0;
        let mut total_machine = 0.0;
        let mut total_emission = 0.0;
        let mut total_research = 0.0;

        // 从road_type中提取数值和单位
        let parts: Vec<&str> = road_type.split_whitespace().collect();
        let road_name = road_type.to_string();
        let unit_str = "m2".to_string();

        self.total_rows.push(ResultRow {
            序号: "1".to_string(),
            编码: "".to_string(),
            名称及规格: road_name.clone(),
            项目名称: road_name,
            单位: unit_str,
            可研估算: "1020".to_string(), // 从图中看到的示例值
            碳排放指数: "".to_string(),
            人工: "0.1116".to_string(),
            材料: "22.1072".to_string(),
            机械: "0.3997".to_string(),
            小计: "22.6185".to_string(),
        });

        // 保存子项目数据
        for row in rows {
            if let Ok(row) = row {
                // 累加各项数据
                if let Ok(labor) = row.人工.parse::<f64>() {
                    total_labor += labor;
                }
                if let Ok(material) = row.材料.parse::<f64>() {
                    total_material += material;
                }
                if let Ok(machine) = row.机械.parse::<f64>() {
                    total_machine += machine;
                }
                if let Ok(research) = row.可研估算.parse::<f64>() {
                    total_research += research;
                }
                self.sub_rows.push(row);
            }
        }

        // 更新道路类型行的合计数据
        if let Some(road_row) = self.total_rows.get_mut(1) {
            total_emission = total_labor + total_material + total_machine;
            road_row.碳排放指数 = format!("{:.2}", total_emission / total_research);
            road_row.人工 = format!("{:.4}", total_labor);
            road_row.材料 = format!("{:.4}", total_material);
            road_row.机械 = format!("{:.4}", total_machine);
            road_row.小计 = format!("{:.4}", total_emission);
        }

        Ok(())
    }
}

impl TableDelegate for ResultTableDelegate {
    fn cols_count(&self, _: &App) -> usize {
        self.columns.len()
    }

    fn rows_count(&self, _: &App) -> usize {
        self.total_rows.len()
    }

    fn col_name(&self, col_ix: usize, _: &App) -> SharedString {
        self.columns[col_ix].clone().into()
    }

    fn col_width(&self, _: usize, _: &App) -> Pixels {
        120.0.into()
    }

    fn col_padding(&self, _: usize, _: &App) -> Option<Edges<Pixels>> {
        Some(Edges::all(px(4.)))
    }

    fn render_th(
        &self,
        col_ix: usize,
        _: &mut Window,
        cx: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        div().child(self.col_name(col_ix, cx))
    }

    fn render_td(
        &self,
        row_ix: usize,
        col_ix: usize,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        let row = &self.total_rows[row_ix];
        let value = match col_ix {
            0 => row.序号.clone(),
            1 => row.项目名称.clone(),
            2 => row.单位.clone(),
            3 => row.可研估算.clone(),
            4 => row.碳排放指数.clone(),
            5 => row.人工.clone(),
            6 => row.材料.clone(),
            7 => row.机械.clone(),
            8 => row.小计.clone(),
            _ => String::new(),
        };

        div().child(value)
    }
}

struct CarbonResultPanel {
    table: Entity<Table<ResultTableDelegate>>,
    focus_handle: FocusHandle,
    story: Entity<ExcelStory>,
    _subscriptions: Vec<Subscription>,
    sub_items_panel: Option<Entity<SubItemsPanel>>,
}

impl CarbonResultPanel {
    pub fn view(
        story: Entity<ExcelStory>,
        window: &mut Window,
        cx: &mut Context<DockArea>,
    ) -> Entity<Self> {
        let delegate = ResultTableDelegate::new();
        let table = cx.new(|cx| Table::new(delegate, window, cx));

        let view = cx.new(|cx| {
            let table_clone = table.clone();
            let subscription = cx.subscribe_in(
                &story,
                window,
                move |this: &mut CarbonResultPanel,
                      story: &Entity<ExcelStory>,
                      _: &UpdateResultTable,
                      window: &mut Window,
                      cx: &mut Context<CarbonResultPanel>| {
                    let story_data = story.read(cx);
                    let project_type = story_data.project_type.clone();
                    let road_type = story_data.road_type.clone();
                    dbg!(&project_type, &road_type);
                    let db_path = story_data.db_path.clone();
                    drop(story_data);

                    table_clone.update(cx, |table, cx| {
                        if project_type.is_empty() || road_type.is_empty() {
                            window.push_notification(
                                Notification::new("请先选择工程类型和道路类型")
                                    .with_type(NotificationType::Warning),
                                cx,
                            );
                        }

                        if let Err(e) =
                            table
                                .delegate_mut()
                                .update_data(&project_type, &road_type, &db_path)
                        {
                            window.push_notification(
                                Notification::new(format!("更新数据失败: {}", e))
                                    .with_type(NotificationType::Error),
                                cx,
                            );
                        }
                        table.refresh(cx);
                    });
                },
            );

            let mut panel = Self {
                table: table.clone(),
                focus_handle: cx.focus_handle(),
                story: story.clone(),
                _subscriptions: vec![subscription],
                sub_items_panel: None,
            };

            // Subscribe to table selection changes
            let table_subscription = cx.subscribe_in(&table, window, {
                let story = story.clone();
                move |this: &mut CarbonResultPanel,
                      _table,
                      event: &TableEvent,
                      window: &mut Window,
                      cx: &mut Context<CarbonResultPanel>| {
                    if let TableEvent::SelectRow(row_ix) = event {
                        if *row_ix == 1 {
                            // 当选中第二行时
                            let story_data = story.read(cx);
                            let dock_area = story_data.dock_area.clone();
                            drop(story_data);

                            // Create sub items panel if not exists
                            if this.sub_items_panel.is_none() {
                                // Create sub items panel
                                let delegate = SubItemsTableDelegate::new();
                                let table = cx.new(|cx| Table::new(delegate, window, cx));
                                let sub_items_panel = cx.new(|cx| SubItemsPanel {
                                    table,
                                    focus_handle: cx.focus_handle(),
                                });

                                // Set sub rows data
                                let table_data = this.table.read(cx).delegate().clone();
                                sub_items_panel.update(cx, |panel, cx| {
                                    panel.table.update(cx, |table, cx| {
                                        table.delegate_mut().set_rows(table_data.sub_rows.clone());
                                        table.refresh(cx);
                                    });
                                });

                                // Add panel to dock area
                                let panel_item = DockItem::tab(
                                    sub_items_panel.clone(),
                                    &Entity::downgrade(&dock_area),
                                    window,
                                    cx,
                                );
                                dock_area.update(cx, |dock_area, cx| {
                                    dock_area.set_right_dock(
                                        panel_item,
                                        Some(px(600.)),
                                        true,
                                        window,
                                        cx,
                                    );
                                });

                                this.sub_items_panel = Some(sub_items_panel);
                            } else if let Some(panel) = &this.sub_items_panel {
                                // Update existing panel data
                                let table_data = this.table.read(cx).delegate().clone();
                                panel.update(cx, |panel, cx| {
                                    panel.table.update(cx, |table, cx| {
                                        table.delegate_mut().set_rows(table_data.sub_rows.clone());
                                        table.refresh(cx);
                                    });
                                });
                            }
                        }
                    }
                }
            });

            panel._subscriptions.push(table_subscription);
            panel
        });

        view
    }
}

impl EventEmitter<PanelEvent> for CarbonResultPanel {}

impl Panel for CarbonResultPanel {
    fn panel_name(&self) -> &'static str {
        "ResultPanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "计算结果".into_any_element()
    }
}

impl Focusable for CarbonResultPanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for CarbonResultPanel {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div().size_full().child(self.table.clone())
    }
}

struct SubItemsPanel {
    table: Entity<Table<SubItemsTableDelegate>>,
    focus_handle: FocusHandle,
}

#[derive(Clone, Default, Debug)]
struct SubItemRow {
    序号: String,
    编码: String,
    名称及规格: String,
    单位: String,
    数量: String,
    碳排放因子: String,
    碳排放量: String,
}

struct SubItemsTableDelegate {
    rows: Vec<SubItemRow>,
    columns: Vec<String>,
}

impl SubItemsTableDelegate {
    fn new() -> Self {
        let columns = vec![
            "序号".to_string(),
            "编码".to_string(),
            "名称及规格".to_string(),
            "单位".to_string(),
            "数量".to_string(),
            "碳排放因子".to_string(),
            "碳排放量".to_string(),
        ];

        Self {
            rows: Vec::new(),
            columns,
        }
    }

    fn set_rows(&mut self, result_rows: Vec<ResultRow>) {
        self.rows = result_rows
            .into_iter()
            .map(|row| {
                let total_emission = if let (Ok(labor), Ok(material), Ok(machine)) = (
                    row.人工.parse::<f64>(),
                    row.材料.parse::<f64>(),
                    row.机械.parse::<f64>(),
                ) {
                    labor + material + machine
                } else {
                    0.0
                };

                let amount = row.可研估算.parse::<f64>().unwrap_or(0.0);
                let factor = if amount > 0.0 {
                    total_emission / amount
                } else {
                    0.0
                };

                SubItemRow {
                    序号: row.序号,
                    编码: row.编码,
                    名称及规格: row.名称及规格,
                    单位: row.单位,
                    数量: row.可研估算,
                    碳排放因子: format!("{:.4}", factor),
                    碳排放量: format!("{:.4}", total_emission),
                }
            })
            .collect();
    }
}

impl SubItemsPanel {
    pub fn view(window: &mut Window, cx: &mut Context<DockArea>) -> Entity<Self> {
        let delegate = SubItemsTableDelegate::new();
        let table = cx.new(|cx| Table::new(delegate, window, cx));

        cx.new(|cx| Self {
            table,
            focus_handle: cx.focus_handle(),
        })
    }
}

impl EventEmitter<PanelEvent> for SubItemsPanel {}

impl Panel for SubItemsPanel {
    fn panel_name(&self) -> &'static str {
        "SubItemsPanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "子项目明细".into_any_element()
    }
}

impl Focusable for SubItemsPanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for SubItemsPanel {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div().size_full().child(self.table.clone())
    }
}

impl TableDelegate for SubItemsTableDelegate {
    fn cols_count(&self, _: &App) -> usize {
        self.columns.len()
    }

    fn rows_count(&self, _: &App) -> usize {
        self.rows.len()
    }

    fn col_name(&self, col_ix: usize, _: &App) -> SharedString {
        self.columns[col_ix].clone().into()
    }

    fn col_width(&self, _: usize, _: &App) -> Pixels {
        120.0.into()
    }

    fn col_padding(&self, _: usize, _: &App) -> Option<Edges<Pixels>> {
        Some(Edges::all(px(4.)))
    }

    fn render_th(
        &self,
        col_ix: usize,
        _: &mut Window,
        cx: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        div().child(self.col_name(col_ix, cx))
    }

    fn render_td(
        &self,
        row_ix: usize,
        col_ix: usize,
        _: &mut Window,
        _: &mut Context<Table<Self>>,
    ) -> impl IntoElement {
        let row = &self.rows[row_ix];
        let value = match col_ix {
            0 => row.序号.clone(),
            1 => row.编码.clone(),
            2 => row.名称及规格.clone(),
            3 => row.单位.clone(),
            4 => row.数量.clone(),
            5 => row.碳排放因子.clone(),
            6 => row.碳排放量.clone(),
            _ => String::new(),
        };

        div().child(value)
    }
}

fn main() {
    let app = Application::new().with_assets(Assets);

    app.run(move |cx| {
        story::init(cx);
        cx.activate(true);
        story::create_new_window("碳排放计算程序", ExcelStory::new_view, cx);
    });
}
