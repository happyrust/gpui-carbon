use gpui::*;
use story::{Assets, Story};
use calamine::{Reader, open_workbook, Xlsx};
use std::collections::HashMap;
use std::ops::Range;
use gpui_component::{
    button::Button, dock::{DockArea, DockPlacement, Panel, PanelView}, h_flex, input::TextInput, label::Label, notification::{Notification, NotificationType}, popup_menu::PopupMenuExt, table::{self, Table, TableDelegate}, v_flex, ContextModal,
};
use std::path::PathBuf;
use gpui::{
    App, Application, Context, Entity, Focusable, FocusHandle,
    IntoElement, Render, Window, SharedString, div, px, Edges,
    impl_actions,
};
use serde::{Serialize, Deserialize};
use schemars::JsonSchema;
use std::sync::Arc;
use gpui::{
    Action,
};

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

impl_actions!(excel_table, [SetProjectType, SetRoadType, ChangeSheet]);

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
        self.rows = data.into_iter().enumerate().map(|(id, data)| ExcelRow { id, data }).collect();
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
    file_path_input: Entity<TextInput>,
    current_sheet: Option<String>,
    required_columns: Vec<String>,
    focus_handle: FocusHandle,
    workbook_data: Option<HashMap<String, (Vec<String>, Vec<HashMap<String, String>>)>>,
    project_type: String,
    road_type: String,
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

    fn new_view(window: &mut Window, cx: &mut App) -> Entity<impl Render + Focusable> {
        Self::view(window, cx)
    }
}

impl Focusable for ExcelStory {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl ExcelStory {
    pub fn view(window: &mut Window, cx: &mut App) -> Entity<Self> {
        cx.new(|cx| Self::new(window, cx))
    }

    fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let focus_handle = cx.focus_handle();
        
        let file_path_input = cx.new(|cx| {
            let mut input = TextInput::new(window, cx);
            input.set_text("assets/excel/指标汇总.xlsx", window, cx);
            input
        });

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
        let dock_area = cx.new(|cx| {
            DockArea::new("excel-table", Some(1), window, cx)
        });

        let story = Self {
            dock_area,
            table,
            file_path_input,
            current_sheet,
            required_columns,
            focus_handle,
            workbook_data: None,
            project_type: "道路工程".to_string(),
            road_type: "主干路 62 m2".to_string(),
        };

        let story_entity = cx.entity();
        // Add panels to dock area
        story.dock_area.update(cx, |dock_area, cx| {
            // Add form panel to the left
            let form_panel = FormPanel::view(story_entity.clone(), window, cx);
            dock_area.add_panel(form_panel, DockPlacement::Left, None, window, cx);

            // Add table panel to the center
            let table_panel = TablePanel::view(story_entity, window, cx);
            dock_area.add_panel(table_panel, DockPlacement::Center, None, window, cx);
        });

        story
    }

    fn change_sheet(&mut self, sheet_name: String, window: &mut Window, cx: &mut Context<Self>) {
        if let Some(workbook_data) = &self.workbook_data {
            if let Some((headers, data)) = workbook_data.get(&sheet_name) {
                let data_len = data.len();
                self.table.update(cx, |table, cx| {
                    table.delegate_mut().set_data(headers.clone(), data.clone());
                    table.refresh(cx);
                });

                window.push_notification(
                    Notification::new(format!("已切换到工作表 '{}', 共 {} 行数据", sheet_name, data_len))
                        .with_type(NotificationType::Success),
                    cx
                );
            }
        }
    }

    fn load_excel(&mut self, window: &mut Window, cx: &mut Context<Self>) {
        let file_path = self.file_path_input.read(cx).text();
        let path = PathBuf::from(file_path.as_ref());

        match open_workbook::<Xlsx<_>, _>(&path) {
            Ok(mut workbook) => {
                let sheet_names = workbook.sheet_names().to_vec();
                let mut workbook_data = HashMap::new();
                
                // Process each sheet
                for sheet_name in &sheet_names {
                    match workbook.worksheet_range(sheet_name) {
                        Ok(range) => {
                            if !range.is_empty() {
                                let rows: Vec<_> = range.rows().collect();
                                
                                // Find header row
                                let mut header_columns = HashMap::new();
                                let header_idx = rows.iter().position(|row| {
                                    header_columns.clear();
                                    for (col_idx, cell) in row.iter().enumerate() {
                                        let cell_value = cell.to_string().trim().to_string();
                                        if !cell_value.is_empty() {
                                            header_columns.insert(col_idx, cell_value);
                                        }
                                    }
                                    
                                    // Check if all required columns are present
                                    self.required_columns.iter().all(|required_col| {
                                        header_columns.values().any(|col| col == required_col)
                                    })
                                });

                                if let Some(header_idx) = header_idx {
                                    // Create a mapping of column index to required column name
                                    let mut column_mapping = HashMap::new();
                                    for (col_idx, header) in header_columns.iter() {
                                        if self.required_columns.contains(header) {
                                            column_mapping.insert(*col_idx, header.clone());
                                        }
                                    }

                                    // Get headers in the required order
                                    let headers = self.required_columns.clone();
                                    
                                    // Convert data
                                    let mut data = Vec::new();
                                    for row in rows.iter().skip(header_idx + 1) {
                                        let mut row_data = HashMap::new();
                                        let mut valid_columns = 0;
                                        
                                        // Process only the mapped columns in the required order
                                        for required_col in &self.required_columns {
                                            if let Some(col_idx) = column_mapping.iter()
                                                .find(|(_, &ref header)| header == required_col)
                                                .map(|(&idx, _)| idx) {
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
                                            data.push(row_data);
                                        }
                                    }

                                    workbook_data.insert(sheet_name.clone(), (headers, data));
                                }
                            }
                        }
                        Err(e) => {
                            window.push_notification(
                                Notification::new(format!("无法读取工作表 '{}': {}", sheet_name, e))
                                    .with_type(NotificationType::Warning),
                                cx
                            );
                        }
                    }
                }

                if workbook_data.is_empty() {
                    window.push_notification(
                        Notification::new("未找到任何有效的工作表数据")
                            .with_type(NotificationType::Error),
                        cx
                    );
                    return;
                }

                // Update sheet select options
                let sheet_options: Vec<_> = sheet_names.into_iter()
                    .filter(|name| workbook_data.contains_key(name))
                    .collect();

                self.current_sheet = sheet_options.first().cloned();

                // Store workbook data
                self.workbook_data = Some(workbook_data);

                // Load first sheet
                if let Some(first_sheet) = sheet_options.first() {
                    self.change_sheet(first_sheet.clone(), window, cx);
                }
            }
            Err(e) => {
                window.push_notification(
                    Notification::new(format!("打开 Excel 文件失败: {}", e))
                        .with_type(NotificationType::Error),
                    cx
                );
            }
        }
    }

    fn render_form(&self, _cx: &mut Context<Self>) -> impl IntoElement {
        let this = self.clone();
        v_flex()
            .gap_4()
            .p_4()
            .min_w(px(350.))
            .bg(hsla(0.0, 0.0, 0.96, 1.0))
            .child(
                v_flex()
                    .gap_2()
                    .child(Label::new("工程类型"))
                    .child(
                        Button::new("project_type")
                            .label(&this.project_type)
                            .popup_menu(move |menu, _, _| {
                                menu.menu("道路工程", Box::new(SetProjectType("道路工程".to_string())))
                                   .menu("交通工程", Box::new(SetProjectType("交通工程".to_string())))
                            }),
                    ),
            )
            .child(
                v_flex()
                    .gap_2()
                    .child(Label::new("道路类型"))
                    .child(
                        Button::new("road_type")
                            .label(&this.road_type)
                            .popup_menu(move |menu, _, _| {
                                menu.menu("主干路 62 m2", Box::new(SetRoadType("主干路 62 m2".to_string())))
                                   .menu("次干路59cm m2", Box::new(SetRoadType("次干路59cm m2".to_string())))
                            }),
                    ),
            )
    }

    fn on_set_project_type(&mut self, action: &SetProjectType, _: &mut Window, cx: &mut Context<Self>) {
        self.project_type = action.0.clone();
        cx.notify();
    }

    fn on_set_road_type(&mut self, action: &SetRoadType, _: &mut Window, cx: &mut Context<Self>) {
        self.road_type = action.0.clone();
        cx.notify();
    }

    fn on_change_sheet(&mut self, action: &ChangeSheet, window: &mut Window, cx: &mut Context<Self>) {
        self.change_sheet(action.sheet_name.clone(), window, cx);
    }
}

impl Render for ExcelStory {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        v_flex()
            .size_full()
            .on_action(cx.listener(Self::on_set_project_type))
            .on_action(cx.listener(Self::on_set_road_type))
            .on_action(cx.listener(Self::on_change_sheet))
            .child(self.dock_area.clone())
    }
}

struct FormPanel {
    story: Entity<ExcelStory>,
    focus_handle: FocusHandle,
}

impl FormPanel {
    pub fn view(story: Entity<ExcelStory>, _window: &mut Window, cx: &mut Context<DockArea>) -> Arc<dyn PanelView> {
        let form_panel = cx.new(|cx| Self {
            story,
            focus_handle: cx.focus_handle(),
        });
        Arc::new(form_panel)
    }

    fn min_width(&self) -> f32 {
        350.0
    }
}

impl Panel for FormPanel {
    fn panel_name(&self) -> &'static str {
        "FormPanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "配置".into_any_element()
    }
}

impl EventEmitter<gpui_component::dock::PanelEvent> for FormPanel {}

impl Focusable for FormPanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for FormPanel {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let this = self.story.read(cx);
        div()
            .size_full()
            .min_w(px(200.))
            .child(
                v_flex()
                    .size_full()
                    .gap_4()
                    .p_4()
                    .bg(hsla(0.0, 0.0, 0.96, 1.0))
                    .child(
                        v_flex()
                            .size_full()
                            .gap_2()
                            .child(Label::new("Excel 文件路径"))
                            .child(
                                div()
                                    .w_full()
                                    .child(this.file_path_input.clone())
                            )
                            .child(
                                h_flex()
                                    .w_full()
                                    .gap_2()
                                    .child(
                                        Button::new("load")
                                            .label("加载 Excel")
                                            .on_click({
                                                let story = self.story.clone();
                                                move |_, window, cx| {
                                                    story.update(cx, |this, cx| {
                                                        this.load_excel(window, cx);
                                                    });
                                                }
                                            })
                                            .flex_grow()
                                    )
                                    .child({
                                        let current_sheet_label = this.current_sheet.as_deref().unwrap_or("选择工作表").to_string();
                                        let button = Button::new("sheet")
                                            .label(current_sheet_label)
                                            .flex_grow();
                                        
                                        if let Some(workbook_data) = &this.workbook_data {
                                            let workbook_data = workbook_data.clone();
                                            button.popup_menu(move |menu, _, _| {
                                                let mut menu = menu;
                                                for sheet_name in workbook_data.keys() {
                                                    let sheet_name = sheet_name.clone();
                                                    menu = menu.menu(
                                                        sheet_name.clone(), 
                                                        Box::new(ChangeSheet {
                                                            sheet_name: sheet_name.clone(),
                                                        })
                                                    );
                                                }
                                                menu
                                            })
                                        } else {
                                            button.popup_menu(|menu, _, _| menu)
                                        }
                                    }),
                            ),
                    )
                    .child(
                        v_flex()
                            .size_full()
                            .gap_2()
                            .child(Label::new("工程类型"))
                            .child(
                                Button::new("project_type")
                                    .label(&this.project_type)
                                    .w_full()
                                    .popup_menu(move |menu, _, _| {
                                        menu.menu("道路工程", Box::new(SetProjectType("道路工程".to_string())))
                                           .menu("交通工程", Box::new(SetProjectType("交通工程".to_string())))
                                    }),
                            ),
                    )
                    .child(
                        v_flex()
                            .size_full()
                            .gap_2()
                            .child(Label::new("道路类型"))
                            .child(
                                Button::new("road_type")
                                    .label(&this.road_type)
                                    .w_full()
                                    .popup_menu(move |menu, _, _| {
                                        menu.menu("主干路 62 m2", Box::new(SetRoadType("主干路 62 m2".to_string())))
                                           .menu("次干路59cm m2", Box::new(SetRoadType("次干路59cm m2".to_string())))
                                    }),
                            ),
                    )
            )
    }
}

struct TablePanel {
    story: Entity<ExcelStory>,
    focus_handle: FocusHandle,
}

impl TablePanel {
    pub fn view(story: Entity<ExcelStory>, _window: &mut Window, cx: &mut Context<DockArea>) -> Arc<dyn PanelView> {
        let panel = cx.new(|cx| Self {
            story,
            focus_handle: cx.focus_handle(),
        });
        Arc::new(panel)
    }
}

impl Panel for TablePanel {
    fn panel_name(&self) -> &'static str {
        "TablePanel"
    }

    fn title(&self, _window: &Window, _cx: &App) -> gpui::AnyElement {
        "数据".into_any_element()
    }
}

impl EventEmitter<gpui_component::dock::PanelEvent> for TablePanel {}

impl Focusable for TablePanel {
    fn focus_handle(&self, _: &App) -> FocusHandle {
        self.focus_handle.clone()
    }
}

impl Render for TablePanel {
    fn render(&mut self, _window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let this = self.story.read(cx);
        div()
            .size_full()
            .child(this.table.clone())
    }
}

fn main() {
    let app = Application::new().with_assets(Assets);

    app.run(move |cx| {
        story::init(cx);
        cx.activate(true);

        story::create_new_window("碳排放计算程序", ExcelStory::view, cx);
    });
} 