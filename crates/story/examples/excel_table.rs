use calamine::{open_workbook, Reader, Xlsx};
use fake::faker::name;
use gpui::{
    div, hsla, impl_actions, px, App, AppContext, Application, BorrowAppContext, Context, Edges,
    Entity, EventEmitter, FocusHandle, Focusable, InteractiveElement, IntoElement, ParentElement,
    Pixels, Render, SharedString, Size, Styled, Subscription, Window,
};
use gpui_component::{
    button::Button, dock::{DockArea, DockItem, DockPlacement, Panel, PanelEvent, PanelView}, dropdown::{Dropdown, DropdownEvent}, h_flex, input::TextInput, label::Label, notification::{Notification, NotificationType}, popup_menu::PopupMenuExt, scroll::ScrollbarShow, table::{self, Table, TableDelegate, TableEvent}, v_flex, ContextModal, Sizable, Theme
};
use rusqlite::{params, Connection, Result as SqliteResult};
use schemars::JsonSchema;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::ops::Range;
use std::path::PathBuf;
use std::sync::Arc;
use story::{debug, Assets, Story};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Deserialize, JsonSchema, Serialize)]
enum Category {
    Labor,    // 人工
    Material, // 材料
    Machine,  // 机械
    None,     // 未分类
}

impl Default for Category {
    fn default() -> Self {
        Category::None
    }
}

impl Category {
    fn from_type_column(名称及规格: &str) -> Option<Self> {
        match 名称及规格.trim() {
            "人工类别" => Some(Category::Labor),
            "材料类别" => Some(Category::Material), 
            "机械类别" => Some(Category::Machine),
            _ => None
        }
    }

    fn prefix(&self) -> &'static str {
        match self {
            Category::Labor => "L",
            Category::Material => "M",
            Category::Machine => "E",
            Category::None => "",
        }
    }

    fn to_string(&self) -> String {
        match self {
            Category::Labor => "labor",
            Category::Material => "material",
            Category::Machine => "machine",
            Category::None => "none",
        }.to_string()
    }

    fn from_string(s: &str) -> Self {
        match s {
            "labor" => Category::Labor,
            "material" => Category::Material,
            "machine" => Category::Machine,
            _ => Category::None,
        }
    }
}

#[derive(Clone)]
struct IndicatorRow {
    id: usize,
    data: HashMap<String, String>,
}

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct ChangeSheet {
    sheet_name: String,
}

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct UpdateResultTable;

#[derive(Clone, Debug, PartialEq, Deserialize, JsonSchema, Serialize)]
pub struct UpdateTypesUIEvent;

impl_actions!(
    indicator_table,
    [
        ChangeSheet,
        UpdateResultTable,
        UpdateTypesUIEvent
    ]
);

struct IndicatorTableDelegate {
    rows: Vec<IndicatorRow>,
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

impl IndicatorTableDelegate {
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
            .map(|(id, data)| IndicatorRow { id, data })
            .collect();
        self.eof = true;
        self.loading = false;
        self.full_loading = false;
    }
}

impl TableDelegate for IndicatorTableDelegate {
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
pub struct IndicatorStory {
    dock_area: Entity<DockArea>,
    table: Entity<Table<IndicatorTableDelegate>>,
    current_sheet: Option<String>,
    required_columns: Vec<String>,
    focus_handle: FocusHandle,
    db_path: String,
}

impl Story for IndicatorStory {
    fn title() -> &'static str {
        "Indicator Table"
    }

    fn description() -> &'static str {
        "A table component that can display indicator data with configurable columns"
    }

    fn klass() -> &'static str {
        "IndicatorStory"
    }

    fn new_view(window: &mut Window, cx: &mut App) -> Entity<Self> {
        cx.new(|cx| Self::new(window, cx))
    }
}

impl Focusable for IndicatorStory {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl EventEmitter<UpdateResultTable> for IndicatorStory {}
impl EventEmitter<UpdateTypesUIEvent> for IndicatorStory {}

impl IndicatorStory {
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

        let delegate = IndicatorTableDelegate::new();
        let table = cx.new(|cx| Table::new(delegate, window, cx));

        // Initialize dock area
        let dock_area = cx.new(|cx| DockArea::new("excel-table", Some(1), window, cx));

        // Create data directory if it doesn't exist
        let data_dir = "data";
        fs::create_dir_all(data_dir).expect("Failed to create data directory");
        let db_path = format!("{}/indicator_data.db", data_dir);

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
                IndicatorTablePanel::view(story_entity.clone(), window, cx),
                &weak_dock_area,
                window,
                cx,
            );

            // Set panels in dock area
            dock_area.set_center(center_panel, window, cx);
            dock_area.set_left_dock(left_panel, Some(px(300.)), true, window, cx);
            dock_area.set_right_dock(right_panel, Some(px(600.)), true, window, cx);
        });

        let indicator_story = Self {
            dock_area,
            table,
            current_sheet,
            required_columns,
            focus_handle,
            db_path,
        };

        // Load resource data
        indicator_story.load_resource_data("assets/excel/人材机数据库.xlsx", window, cx);

        indicator_story
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
                category TEXT,  -- 添加category字段
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
             AND (编码 IS NOT NULL OR 名称及规格 LIKE '%类别')  -- 修改过滤条件，确保能读取到类别标题行
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

                        // Clear existing data before importing new data
                        tx.execute("DELETE FROM excel_data", []).unwrap();
                        tx.execute("DELETE FROM sheets", []).unwrap();

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

                                        // Track current category and counts
                                        let mut current_category = Category::None;
                                        let mut category_counts = HashMap::new();
                                        category_counts.insert(Category::Labor, 0);
                                        category_counts.insert(Category::Material, 0);
                                        category_counts.insert(Category::Machine, 0);
                                        category_counts.insert(Category::None, 0);

                                        debug!("Starting to process rows for sheet:", sheet_name);

                                        // Insert data
                                        for row in rows.iter().skip(header_idx + 1) {
                                            let mut row_data = HashMap::new();
                                            let mut valid_columns = 0;

                                            // 获取名称及规格字段
                                            let 名称及规格 = row.get(2).map(|c| c.to_string()).unwrap_or_default();

                                            // 检查是否是类别标题行
                                            if 名称及规格.trim() == "人工类别" {
                                                current_category = Category::Labor;
                                                debug!("Found labor category");
                                                continue;
                                            } else if 名称及规格.trim() == "材料类别" {
                                                current_category = Category::Material;
                                                debug!("Found material category");
                                                continue;
                                            } else if 名称及规格.trim() == "机械类别" {
                                                current_category = Category::Machine;
                                                debug!("Found machine category");
                                                continue;
                                            }

                                            // 收集行数据
                                            for required_col in &self.required_columns {
                                                if let Some(col_idx) = column_mapping
                                                    .iter()
                                                    .find(|(_, header)| *header == required_col)
                                                    .map(|(&idx, _)| idx)
                                                {
                                                    if let Some(cell) = row.get(col_idx) {
                                                        let cell_value = cell.to_string().trim().to_string();
                                                        if !cell_value.is_empty() {
                                                            row_data.insert(required_col.clone(), cell_value);
                                                            valid_columns += 1;
                                                        }
                                                    }
                                                }
                                            }

                                            if valid_columns > 0 {
                                                // 更新类别计数
                                                if let Some(count) = category_counts.get_mut(&current_category) {
                                                    *count += 1;
                                                }

                                                // Modify 编码 based on category if it doesn't already have a prefix
                                                if let Some(code) = row_data.get_mut("编码") {
                                                    if !code.starts_with('L') && !code.starts_with('M') && !code.starts_with('E') {
                                                        *code = format!("{}{}", current_category.prefix(), code);
                                                        debug!("Added prefix to code:", code, "Category:", &current_category);
                                                    }
                                                }

                                                tx.execute(
                                                    "INSERT INTO excel_data (
                                                        sheet_id, 序号, 编码, 名称及规格, 单位, 数量, 市场价, 合计, category
                                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                                    params![
                                                        sheet_id,
                                                        row_data.get("序号").unwrap_or(&String::new()),
                                                        row_data.get("编码").unwrap_or(&String::new()),
                                                        row_data.get("名称及规格").unwrap_or(&String::new()),
                                                        row_data.get("单位").unwrap_or(&String::new()),
                                                        row_data.get("数量").unwrap_or(&String::new()),
                                                        row_data.get("市场价").unwrap_or(&String::new()),
                                                        row_data.get("合计").unwrap_or(&String::new()),
                                                        current_category.to_string(),
                                                    ],
                                                ).unwrap();
                                            }
                                        }

                                        // Print category statistics
                                        debug!("Category counts for sheet:", sheet_name);
                                        debug!("Labor items:", category_counts[&Category::Labor]);
                                        debug!("Material items:", category_counts[&Category::Material]);
                                        debug!("Machine items:", category_counts[&Category::Machine]);
                                        debug!("Uncategorized items:", category_counts[&Category::None]);

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

                            // Load first sheet
                            if let Some(first_sheet) = sheet_names.first() {
                                self.current_sheet = Some(first_sheet.clone());
                                self.change_sheet(first_sheet.clone(), window, cx);
                            }

                            window.push_notification(
                                Notification::new("指标数据已成功导入到数据库")
                                    .with_type(NotificationType::Success),
                                cx,
                            );

                            // 直接触发结果表更新
                            cx.emit(UpdateResultTable);
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

    fn load_resource_data(&self, file_path: &str, window: &mut Window, cx: &mut Context<Self>) {
        let path = PathBuf::from(file_path);

        match open_workbook::<Xlsx<_>, _>(&path) {
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
                            // debug!(sheet_name, table_name);
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
                                        // debug!(&column_indices);
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
                        debug!("Error connecting to database: {}", &e);
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

impl Render for IndicatorStory {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        v_flex()
            .size_full()
            .on_action(cx.listener(Self::on_change_sheet))
            .child(self.dock_area.clone())
    }
}

struct ParamFormPane {
    story: Entity<IndicatorStory>,
    focus_handle: FocusHandle,
    file_path_input: Entity<TextInput>,
    carbon_ref_input: Entity<TextInput>,
}

impl ParamFormPane {
    pub fn view(
        story: Entity<IndicatorStory>,
        window: &mut Window,
        cx: &mut Context<DockArea>,
    ) -> Entity<Self> {
        let file_path_input = cx.new(|cx| {
            let mut input = TextInput::new(window, cx);
            input.set_text("assets/excel/指标汇总.xlsx", window, cx);
            input
        });

        let carbon_ref_input = cx.new(|cx| {
            let mut input = TextInput::new(window, cx);
            input.set_text("assets/excel/人材机数据库.xlsx", window, cx);
            input
        });

        cx.new(|cx| Self {
            story: story.clone(),
            focus_handle: cx.focus_handle(),
            file_path_input: file_path_input.clone(),
            carbon_ref_input: carbon_ref_input.clone(),
        })
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
        let file_path_input = self.file_path_input.clone();
        let carbon_ref_input = self.carbon_ref_input.clone();
        div()
            .flex()
            .flex_col()
            .w_auto()
            .p_4()
            .bg(hsla(0.0, 0.0, 0.96, 1.0))
            .child(
                v_flex()
                    .gap_4()
                    .child(
                        v_flex()
                            .gap_2()
                            .child(Label::new("单位指标文件"))
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
                                                                .set_title("选择单位指标文件")
                                                                .pick_file()
                                                                .await
                                                        {
                                                            let path_str = path
                                                                .path()
                                                                .to_string_lossy()
                                                                .to_string();
                                                            awc.update(|win, cx| {
                                                                file_path_input.update(
                                                                    cx,
                                                                    |input, cx| {
                                                                        input.set_text(
                                                                            &path_str, win, cx,
                                                                        );
                                                                    },
                                                                );
                                                            }).unwrap();
                                                        }
                                                    })
                                                    .detach();
                                            }
                                        })),
                                ),
                            )
                            .child(
                                Button::new("load")
                                    .label("导入单位指标")
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
                                    .w_full(),
                            )
                            .child(
                                v_flex()
                                    .gap_2()
                                    .child(Label::new("人材机数据文件"))
                                    .child(
                                        div().w_full().child(
                                            h_flex()
                                                .gap_1()
                                                .child(div().flex_grow().child(carbon_ref_input.clone()))
                                                .child(Button::new("select-carbon-ref").label("...").on_click({
                                                    let carbon_ref_input = carbon_ref_input.clone();
                                                    move |_, window, cx| {
                                                        let carbon_ref_input = carbon_ref_input.clone();
                                                        window
                                                            .spawn(cx, |mut awc| async move {
                                                                if let Some(path) =
                                                                    rfd::AsyncFileDialog::new()
                                                                        .add_filter(
                                                                            "Excel files",
                                                                            &["xlsx"],
                                                                        )
                                                                        .set_title("选择人材机数据文件")
                                                                        .pick_file()
                                                                        .await
                                                                {
                                                                    let path_str = path
                                                                        .path()
                                                                        .to_string_lossy()
                                                                        .to_string();
                                                                    awc.update(|win, cx| {
                                                                        carbon_ref_input.update(
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
                                        Button::new("load-carbon-ref")
                                            .label("导入人材机数据")
                                            .on_click({
                                                let story = self.story.clone();
                                                let carbon_ref_input = self.carbon_ref_input.clone();
                                                move |_, window, cx| {
                                                    let file_path = carbon_ref_input.read(cx).text();
                                                    story.update(cx, |this, cx| {
                                                        this.load_resource_data(&file_path, window, cx);
                                                    });
                                                }
                                            })
                                            .w_full(),
                                    ),
                            ),
                    )
            )
    }
}

struct IndicatorTablePanel {
    story: Entity<IndicatorStory>,
    focus_handle: FocusHandle,
}

impl IndicatorTablePanel {
    pub fn view(
        story: Entity<IndicatorStory>,
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

impl EventEmitter<PanelEvent> for IndicatorTablePanel {}

impl Panel for IndicatorTablePanel {
    fn panel_name(&self) -> &'static str {
        "TablePanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "数据".into_any_element()
    }
}

impl Focusable for IndicatorTablePanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for IndicatorTablePanel {
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

#[derive(Clone, Default, Debug)]
struct SubItemRow {
    序号: String,
    编码: String,
    名称及规格: String,
    单位: String,
    数量: String,
    市场价: String,
    合计: String,
    category: Category,  // 添加类别字段
}

#[derive(Clone)]
struct ResultTableDelegate {
    total_rows: Vec<ResultRow>,
    sub_rows: Vec<SubItemRow>,
    columns: Vec<String>,
}

impl ResultTableDelegate {
    fn new() -> Self {
        let columns = vec![
            "序号".to_string(),
            "项目名称".to_string(),
            "单位".to_string(),
            "工程量".to_string(),  // 改名：可研估算 -> 工程量
            "总碳排放量".to_string(),  // 改名：碳排放指数 -> 总碳排放量
            "主要人材机耗量".to_string(),  // 新增列
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

    fn update_data(&mut self, db_path: &str) -> SqliteResult<()> {
        self.total_rows.clear();
        self.sub_rows.clear();

        let conn = Connection::open(db_path)?;

        // Debug: Check if tables exist and have data
        debug!("Checking database tables");
        let tables = ["sheets", "excel_data", "labor", "material", "machine"];
        for table in &tables {
            let count: i64 = conn.query_row(
                &format!("SELECT COUNT(*) FROM {}", table),
                [],
                |row| row.get(0),
            )?;
            debug!("Table {} has {} rows", table, count);
        }

        // 获取所有工作表并按类型分组
        let mut sheets_stmt = conn.prepare(
            "SELECT id, name, 
             -- Get the type from the first part of the sheet name before a space
             CASE 
                WHEN instr(name, ' ') > 0 THEN substr(name, 1, instr(name, ' ')-1)
                ELSE name
             END as type,
             -- Extract the specific name part (everything after the first space)
             CASE
                WHEN instr(name, ' ') > 0 THEN substr(name, instr(name, ' ')+1)
                ELSE ''
             END as short_name
             FROM sheets 
             ORDER BY type, id"
        )?;

        let sheets = sheets_stmt.query_map([], |row| {
            Ok((
                row.get::<_, i64>(0)?,
                row.get::<_, String>(1)?,
                row.get::<_, String>(2)?,
                row.get::<_, String>(3)?
            ))
        })?;

        let mut current_type = String::new();
        let mut type_index = 1;

        // Collect all sheet results
        let mut sheet_results = Vec::new();
        for sheet_result in sheets {
            sheet_results.push(sheet_result?);
        }
        
        debug!("Found {} sheets in database", sheet_results.len());
        if sheet_results.is_empty() {
            debug!("No sheets found in database. Check if data was imported correctly.");
            return Ok(());
        }

        for (sheet_id, sheet_name, sheet_type, short_name) in sheet_results {
            debug!("Processing sheet: {} (id: {}, type: {})", &sheet_name, sheet_id, &sheet_type);
            
            // 如果是新的工程类型，添加类型标题行
            if current_type != sheet_type {
                current_type = sheet_type.clone();
                self.total_rows.push(ResultRow {
                    序号: match type_index {
                        1 => "一".to_string(),
                        2 => "二".to_string(),
                        3 => "三".to_string(),
                        _ => format!("{}", type_index),
                    },
                    编码: "".to_string(),
                    名称及规格: "".to_string(),
                    项目名称: current_type.clone(),
                    单位: "".to_string(),
                    可研估算: "".to_string(),
                    碳排放指数: "".to_string(),
                    人工: "".to_string(),
                    材料: "".to_string(),
                    机械: "".to_string(),
                    小计: "".to_string(),
                });
                type_index += 1;
            }

            // 首先检查该工作表是否有数据
            let data_count: i64 = conn.query_row(
                "SELECT COUNT(*) FROM excel_data WHERE sheet_id = ?",
                params![sheet_id],
                |row| row.get(0),
            )?;
            debug!("Sheet {} has {} data rows", &sheet_name, data_count);
            
            // 检查分类数据
            let category_count: Vec<(String, i64)> = {
                let mut stmt = conn.prepare(
                    "SELECT category, COUNT(*) FROM excel_data 
                     WHERE sheet_id = ? AND category IS NOT NULL 
                     GROUP BY category"
                )?;
                let results = stmt.query_map(params![sheet_id], |row| {
                    Ok((row.get::<_, String>(0)?, row.get::<_, i64>(1)?))
                })?;
                let mut counts = Vec::new();
                for result in results {
                    counts.push(result?);
                }
                counts
            };
            
            for (category, count) in &category_count {
                debug!("Category '{}' has {} items", category, count);
            }

            // 查询该工作表的分类汇总数据
            let mut stmt = conn.prepare(
                "WITH combined_data AS (
                    SELECT 
                        e.*,
                        CASE 
                            WHEN e.编码 LIKE 'L%' THEN l.carbon_factor
                            WHEN e.编码 LIKE 'M%' THEN m.carbon_factor
                            WHEN e.编码 LIKE 'E%' THEN mc.carbon_factor
                            ELSE 1.0
                        END as carbon_factor
                    FROM excel_data e
                    LEFT JOIN labor l ON substr(e.编码, 2) = l.code
                    LEFT JOIN material m ON substr(e.编码, 2) = m.code
                    LEFT JOIN machine mc ON substr(e.编码, 2) = mc.code
                    WHERE e.sheet_id = ?
                )
                SELECT 
                    category,
                    SUM(CAST(NULLIF(数量,'') AS FLOAT) * carbon_factor) as total_emission
                FROM combined_data
                WHERE category IS NOT NULL AND 数量 != ''
                GROUP BY category"
            )?;

            let mut total_labor = 0.0;
            let mut total_material = 0.0;
            let mut total_machine = 0.0;
            let category_totals = stmt.query_map(params![sheet_id], |row| {
                let category: String = row.get(0)?;
                let total: f64 = row.get(1)?;
                debug!("Category: {}, Total: {}", &category, total);
                match category.as_str() {
                    "labor" => total_labor = total,
                    "material" => total_material = total,
                    "machine" => total_machine = total,
                    _ => (),
                }
                Ok(())
            })?;

            // 消费迭代器
            for result in category_totals {
                result?; // Process each result to ensure the query runs
            }

            debug!("Totals for sheet {}: Labor: {}, Material: {}, Machine: {}", 
                   &sheet_name, total_labor, total_material, total_machine);

            let total_emission = total_labor + total_material + total_machine;

            // 添加工作表数据行
            self.total_rows.push(ResultRow {
                序号: format!("{}", self.total_rows.len() + 1),
                编码: "".to_string(),
                名称及规格: "".to_string(),
                项目名称: short_name.trim().to_string(),
                单位: "m2".to_string(),
                可研估算: "用户填写".to_string(),  // 工程量，等待用户填写
                碳排放指数: format!("D{}*I{}", self.total_rows.len() + 1, self.total_rows.len() + 1),  // 总碳排放量公式
                人工: format!("{:.2}", total_labor),
                材料: format!("{:.2}", total_material),
                机械: format!("{:.2}", total_machine),
                小计: format!("{:.2}", total_emission),
            });

            // 加载子项目明细数据用于详情面板
            let mut sub_stmt = conn.prepare(
                "SELECT 序号, 编码, 名称及规格, 单位, 数量, 市场价, 合计, category
                 FROM excel_data 
                 WHERE sheet_id = ?
                 AND (编码 IS NOT NULL OR 名称及规格 LIKE '%类别')
                 ORDER BY id"
            )?;
            
            let sub_rows = sub_stmt.query_map(params![sheet_id], |row| {
                Ok(SubItemRow {
                    序号: row.get(0)?,
                    编码: row.get(1)?,
                    名称及规格: row.get(2)?,
                    单位: row.get(3)?,
                    数量: row.get(4)?,
                    市场价: row.get(5)?,
                    合计: row.get(6)?,
                    category: Category::from_string(&row.get::<_, String>(7).unwrap_or_default()),
                })
            })?;
            
            for sub_row in sub_rows {
                self.sub_rows.push(sub_row?);
            }
            
            debug!("Added {} sub items for sheet {}", self.sub_rows.len(), &sheet_name);
        }

        debug!("Final results: {} total rows, {} sub rows", self.total_rows.len(), self.sub_rows.len());
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
        
        // Check if this is a category row - now more dynamic
        // It's a category row if:
        // 1. The name doesn't contain spaces (single word)
        // 2. Any field other than 序号 and 项目名称 is empty
        let is_category_row = !row.项目名称.contains(' ') && row.单位.is_empty() && row.碳排放指数.is_empty();
        
        let value = match col_ix {
            0 => row.序号.clone(),
            1 => row.项目名称.clone(), // Already correctly formatted from the SQL query
            2 => if is_category_row { String::new() } else { row.单位.clone() },
            3 => if is_category_row { String::new() } else { row.可研估算.clone() },
            4 => if is_category_row { String::new() } else { row.碳排放指数.clone() },
            5 => if is_category_row { String::new() } else { String::new() }, // 主要人材机耗量列留空
            6 => if is_category_row { String::new() } else { row.人工.clone() },
            7 => if is_category_row { String::new() } else { row.材料.clone() },
            8 => if is_category_row { String::new() } else { row.机械.clone() },
            9 => if is_category_row { String::new() } else { row.小计.clone() },
            _ => String::new(),
        };

        div().child(value)
    }
}

struct CarbonResultPanel {
    table: Entity<Table<ResultTableDelegate>>, // 表格实体
    focus_handle: FocusHandle, // 焦点句柄
    story: Entity<IndicatorStory>, // Excel 故事实体
    _subscriptions: Vec<Subscription>, // 订阅列表
    sub_items_panel: Option<Entity<ResultDetailsPanel>>, // 子项目面板
}

impl CarbonResultPanel {
    pub fn view(
        story: Entity<IndicatorStory>, // Excel 故事实体
        window: &mut Window, // 窗口引用
        cx: &mut Context<DockArea>, // 上下文
    ) -> Entity<Self> {
        let delegate = ResultTableDelegate::new(); // 创建结果表委托
        let table = cx.new(|cx| Table::new(delegate, window, cx)); // 创建新表格

        let view = cx.new(|cx| {
            let table_clone = table.clone(); // 克隆表格
            let subscription = cx.subscribe_in(
                &story,
                window,
                move |this: &mut CarbonResultPanel,
                      story: &Entity<IndicatorStory>,
                      _: &UpdateResultTable,
                      window: &mut Window,
                      cx: &mut Context<CarbonResultPanel>| {
                    let story_data = story.read(cx); // 读取故事数据
                    let db_path = story_data.db_path.clone(); // 数据库路径
                    drop(story_data); // 释放故事数据

                    table_clone.update(cx, |table, cx| {
                        if let Err(e) =
                            table
                                .delegate_mut()
                                .update_data(&db_path)
                        {
                            window.push_notification(
                                Notification::new(format!("更新数据失败: {}", e))
                                    .with_type(NotificationType::Error),
                                cx,
                            );
                        }
                        table.refresh(cx); // 刷新表格
                    });
                },
            );

            let mut panel = Self {
                table: table.clone(), // 克隆表格
                focus_handle: cx.focus_handle(), // 获取焦点句柄
                story: story.clone(), // 克隆故事
                _subscriptions: vec![subscription], // 添加订阅
                sub_items_panel: None, // 初始化子项目面板
            };

            // 订阅表格选择更改事件
            let table_subscription = cx.subscribe_in(&table, window, {
                let story = story.clone(); // 克隆故事
                move |this: &mut CarbonResultPanel,
                      _table,
                      event: &TableEvent,
                      window: &mut Window,
                      cx: &mut Context<CarbonResultPanel>| {
                    if let TableEvent::SelectRow(row_ix) = event {
                        debug!("Row selected: {}", row_ix); // 调试信息
                        
                        // 忽略类型标题行，它们的项目名称是工程类型名称
                        if let Some(row) = this.table.read(cx).delegate().total_rows.get(*row_ix) {
                            debug!("Selected row project name: {}", &row.项目名称);
                            
                            // 如果是工程类型行（只有类型名，没有空格），不显示子项
                            let is_category_row = !row.项目名称.contains(' ') && row.单位.is_empty();
                            if is_category_row {
                                debug!("Skipping category row");
                                return;
                            }
                            
                            let story_data = story.read(cx); // 读取故事数据
                            let dock_area = story_data.dock_area.clone(); // 克隆停靠区域
                            let sheet_name = row.项目名称.clone(); // 获取选中行的工作表名称
                            drop(story_data); // 释放故事数据

                            // 过滤子项目数据 - 目前所有子项目都在sub_rows中，但实际需要按工作表筛选
                            // 这里先创建面板，之后再通过数据库查询获取该工作表的详细数据
                            debug!("Showing sub items for sheet: {}", &sheet_name);
                            
                            if this.sub_items_panel.is_none() {
                                debug!("Creating new sub items panel");
                                // 创建子项目面板
                                let delegate = ResultDetailsTableDelegate::new();
                                let table = cx.new(|cx| Table::new(delegate, window, cx));
                                let sub_items_panel = cx.new(|cx| ResultDetailsPanel {
                                    table,
                                    focus_handle: cx.focus_handle(),
                                });
                                
                                // 将面板添加到停靠区域
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
                            }
                            
                            // 从数据库加载该工作表的详细数据
                            if let Some(panel) = &this.sub_items_panel {
                                this.load_sub_items_for_sheet(&sheet_name, panel, window, cx);
                            }
                        }
                    }
                }
            });

            panel._subscriptions.push(table_subscription); // 添加表格订阅
            panel
        });

        view
    }

    fn load_sub_items_for_sheet(
        &self,
        sheet_name: &str, 
        panel: &Entity<ResultDetailsPanel>,
        window: &mut Window,
        cx: &mut Context<CarbonResultPanel>
    ) {
        debug!("Loading sub items for sheet: {}", sheet_name);
        
        // Get database path from story
        let db_path = self.story.read(cx).db_path.clone();
        
        // Connect to database
        match Connection::open(&db_path) {
            Ok(conn) => {
                // First get sheet_id - need to look for sheets containing the short_name
                // We'll first need to get all sheet names to find the right one
                let sheet_id = match conn.prepare("SELECT id, name FROM sheets") {
                    Ok(mut stmt) => {
                        let mut matched_id = None;
                        
                        let rows = match stmt.query_map([], |row| {
                            Ok((row.get::<_, i64>(0)?, row.get::<_, String>(1)?))
                        }) {
                            Ok(rows) => rows,
                            Err(e) => {
                                debug!("Error querying sheet names: {}", &e);
                                return;
                            }
                        };
                        
                        // Find sheet with this specific short name
                        for row_result in rows {
                            if let Ok((id, name)) = row_result {
                                if let Some(pos) = name.find(' ') {
                                    let sheet_short_name = &name[pos+1..];
                                    if sheet_short_name == sheet_name {
                                        matched_id = Some(id);
                                        break;
                                    }
                                }
                            }
                        }
                        
                        matched_id
                    },
                    Err(e) => {
                        debug!("Error preparing sheet query: {}", &e);
                        return;
                    }
                };

                // Use the sheet_id if found
                match sheet_id {
                    Some(sheet_id) => {
                        debug!("Found sheet_id: {} for {}", sheet_id, sheet_name);
                        
                        // First ensure we have data for this sheet - help debug issues
                        let count: i64 = match conn.query_row(
                            "SELECT COUNT(*) FROM excel_data WHERE sheet_id = ?",
                            params![sheet_id],
                            |row| row.get(0)
                        ) {
                            Ok(count) => count,
                            Err(e) => {
                                debug!("Error checking row count: {}", &e);
                                0
                            }
                        };
                        
                        if count == 0 {
                            debug!("No data found for sheet: {}", sheet_name);
                            window.push_notification(
                                Notification::new(format!("工作表 {} 没有数据", sheet_name))
                                    .with_type(NotificationType::Warning),
                                cx,
                            );
                            return;
                        }
                        
                        debug!("Sheet {} has {} rows of data", sheet_name, count);
                        
                        // Query sub items for this sheet
                        let mut sub_items = Vec::new();
                        
                        // Check all categories present
                        let categories: Vec<String> = {
                            let mut stmt = match conn.prepare(
                                "SELECT DISTINCT category FROM excel_data 
                                 WHERE sheet_id = ? AND category IS NOT NULL"
                            ) {
                                Ok(stmt) => stmt,
                                Err(e) => {
                                    debug!("Error preparing category query: {}", &e);
                                    return;
                                }
                            };
                            
                            let rows = match stmt.query_map(params![sheet_id], |row| row.get(0)) {
                                Ok(rows) => rows,
                                Err(e) => {
                                    debug!("Error querying categories: {}", &e);
                                    return;
                                }
                            };
                            
                            let mut result = Vec::new();
                            for row in rows {
                                match row {
                                    Ok(category) => result.push(category),
                                    Err(e) => {
                                        debug!("Error reading category: {}", &e);
                                        continue;
                                    }
                                }
                            }
                            result
                        };
                        
                        debug!("Sheet has the following categories: {:?}", &categories);
                        
                        // Process each main category one at a time (labor, material, machine)
                        for category_name in ["labor", "material", "machine"] {
                            // Skip if this category doesn't exist in this sheet
                            if !categories.contains(&category_name.to_string()) {
                                debug!("Sheet doesn't contain category: {}", category_name);
                                continue;
                            }
                            
                            // Create a category header row
                            let header_text = match category_name {
                                "labor" => "人工类别",
                                "material" => "材料类别",
                                "machine" => "机械类别",
                                _ => "未知类别"
                            };
                            
                            // Add category header
                            let header = SubItemRow {
                                序号: "".to_string(),
                                编码: "".to_string(),
                                名称及规格: header_text.to_string(),
                                单位: "".to_string(),
                                数量: "".to_string(),
                                市场价: "".to_string(),
                                合计: "".to_string(),
                                category: Category::from_string(category_name),
                            };
                            
                            sub_items.push(header);
                            
                            // Get items for this category
                            match conn.prepare(
                                "SELECT 序号, 编码, 名称及规格, 单位, 数量, 市场价, 合计 
                                 FROM excel_data 
                                 WHERE sheet_id = ? AND category = ? AND 名称及规格 NOT LIKE '%类别'
                                 ORDER BY id"
                            ) {
                                Ok(mut stmt) => {
                                    match stmt.query_map(params![sheet_id, category_name], |row| {
                                        Ok(SubItemRow {
                                            序号: row.get(0).unwrap_or_default(),
                                            编码: row.get(1).unwrap_or_default(),
                                            名称及规格: row.get(2).unwrap_or_default(),
                                            单位: row.get(3).unwrap_or_default(),
                                            数量: row.get(4).unwrap_or_default(),
                                            市场价: row.get(5).unwrap_or_default(),
                                            合计: row.get(6).unwrap_or_default(),
                                            category: Category::from_string(category_name),
                                        })
                                    }) {
                                        Ok(rows) => {
                                            let mut items = Vec::new();
                                            for (i, result) in rows.enumerate() {
                                                match result {
                                                    Ok(mut row) => {
                                                        // Update row number for better display
                                                        row.序号 = (i + 1).to_string();
                                                        items.push(row);
                                                    },
                                                    Err(e) => {
                                                        debug!("Error reading category item: {}", &e);
                                                        continue;
                                                    }
                                                }
                                            }
                                            
                                            debug!("Added {} items for category {}", items.len(), category_name);
                                            
                                            // Add all items for this category
                                            sub_items.extend(items);
                                            
                                            // Add an empty row as separator if not the last category
                                            if category_name != "machine" {
                                                sub_items.push(SubItemRow::default());
                                            }
                                        },
                                        Err(e) => {
                                            debug!("Error querying category items: {}", &e);
                                            continue;
                                        }
                                    }
                                },
                                Err(e) => {
                                    debug!("Error preparing category items query: {}", &e);
                                    continue;
                                }
                            }
                        }
                        
                        debug!("Total sub items: {}", sub_items.len());
                        
                        if sub_items.is_empty() {
                            window.push_notification(
                                Notification::new(format!("工作表 {} 没有可显示的子项", sheet_name))
                                    .with_type(NotificationType::Warning),
                                cx,
                            );
                            return;
                        }
                        
                        // Update the panel with the sub items
                        panel.update(cx, |panel, cx| {
                            panel.table.update(cx, |table, cx| {
                                table.delegate_mut().set_rows(sub_items);
                                table.refresh(cx);
                            });
                        });
                        
                        // Show the panel if it was previously hidden
                        window.push_notification(
                            Notification::new(format!("已显示工作表 {} 的子项目详情", sheet_name))
                                .with_type(NotificationType::Info),
                            cx,
                        );
                    },
                    None => {
                        debug!("No matching sheet found for {}", sheet_name);
                        window.push_notification(
                            Notification::new(format!("未找到匹配的工作表: {}", sheet_name))
                                .with_type(NotificationType::Error),
                            cx,
                        );
                    }
                }
            },
            Err(e) => {
                debug!("Error connecting to database: {}", &e);
                window.push_notification(
                    Notification::new(format!("数据库连接失败: {}", e))
                        .with_type(NotificationType::Error),
                    cx,
                );
            }
        }
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

struct ResultDetailsPanel {
    table: Entity<Table<ResultDetailsTableDelegate>>, // 表格实体
    focus_handle: FocusHandle, // 焦点句柄
}

struct ResultDetailsTableDelegate {
    rows: Vec<SubItemRow>,
    columns: Vec<String>,
}

impl ResultDetailsTableDelegate {
    fn new() -> Self {
        let columns = vec![
            "序号".to_string(),
            "编码".to_string(),
            "名称及规格".to_string(),
            "单位".to_string(),
            "数量".to_string(),
            "市场价".to_string(),
            "合计".to_string(),
        ];

        Self {
            rows: Vec::new(),
            columns,
        }
    }

    fn set_rows(&mut self, result_rows: Vec<SubItemRow>) {
        debug!("Setting rows in SubItemsTableDelegate, received rows: {}", result_rows.len());
        
        // 清空现有数据
        self.rows.clear();

        // 当前类别
        let mut _current_category = Category::None;
        let mut current_rows: Vec<SubItemRow> = Vec::new();

        // 遍历所有行
        for row in result_rows {
            // 检查是否是类别标题行
            if let Some(category) = Category::from_type_column(&row.名称及规格) {
                // 如果有之前的类别数据，先添加到结果中
                if !current_rows.is_empty() {
                    // 添加当前类别的所有行
                    for (i, mut row) in current_rows.drain(..).enumerate() {
                        row.序号 = (i + 1).to_string();
                        self.rows.push(row);
                    }
                    // 添加空行分隔
                    self.rows.push(SubItemRow::default());
                }

                // 更新当前类别
                _current_category = category;
                // 添加类别标题行
                self.rows.push(row);
            } else if !row.编码.is_empty() {
                // 普通数据行，添加到当前类别的行集合中
                current_rows.push(row);
            }
        }

        // 处理最后一个类别的数据
        if !current_rows.is_empty() {
            for (i, mut row) in current_rows.drain(..).enumerate() {
                row.序号 = (i + 1).to_string();
                self.rows.push(row);
            }
        }

        debug!("Final number of rows in delegate: {}", self.rows.len());
    }
}

impl ResultDetailsPanel {
    pub fn view(window: &mut Window, cx: &mut Context<DockArea>) -> Entity<Self> {
        let delegate = ResultDetailsTableDelegate::new();
        let table = cx.new(|cx| Table::new(delegate, window, cx));

        cx.new(|cx| Self {
            table,
            focus_handle: cx.focus_handle(),
        })
    }
}

impl EventEmitter<PanelEvent> for ResultDetailsPanel {}

impl Panel for ResultDetailsPanel {
    fn panel_name(&self) -> &'static str {
        "SubItemsPanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "子项目明细".into_any_element()
    }
}

impl Focusable for ResultDetailsPanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for ResultDetailsPanel {
    fn render(&mut self, _window: &mut Window, _cx: &mut Context<Self>) -> impl IntoElement {
        div().size_full().child(self.table.clone())
    }
}

impl TableDelegate for ResultDetailsTableDelegate {
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
            5 => row.市场价.clone(),
            6 => row.合计.clone(),
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
        story::create_new_window("碳排放计算程序", IndicatorStory::new_view, cx);
    });
}


