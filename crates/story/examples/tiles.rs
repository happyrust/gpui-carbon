use anyhow::{Context as _, Result};
use gpui::*;
use std::time::Duration;
use story::{Assets, ButtonStory, IconStory, StoryContainer};
use ui::{
    dock::{DockArea, DockAreaState, DockEvent, DockItem},
    ActiveTheme, Root, TitleBar,
};

actions!(main_menu, [Quit]);

const TILES_DOCK_AREA: DockAreaTab = DockAreaTab {
    id: "story-tiles",
    version: 1,
};

actions!(workspace, [Open, CloseWindow]);

pub fn init(cx: &mut App) {
    cx.on_action(|_action: &Open, _cx: &mut App| {});

    ui::init(cx);
    story::init(cx);
}

pub struct StoryTiles {
    dock_area: Entity<DockArea>,
    last_layout_state: Option<DockAreaState>,
    _save_layout_task: Option<Task<()>>,
}

struct DockAreaTab {
    id: &'static str,
    version: usize,
}

impl StoryTiles {
    pub fn new(window: &mut Window, cx: &mut Context<Self>) -> Self {
        let dock_area = cx.new(|cx| {
            DockArea::new(
                TILES_DOCK_AREA.id,
                Some(TILES_DOCK_AREA.version),
                window,
                cx,
            )
        });
        let weak_dock_area = dock_area.downgrade();

        match Self::load_tiles(dock_area.clone(), window, cx) {
            Ok(_) => {
                println!("load tiles success");
            }
            Err(err) => {
                eprintln!("load tiles error: {:?}", err);
                Self::reset_default_layout(weak_dock_area, window, cx);
            }
        };

        cx.subscribe_in(
            &dock_area,
            window,
            |this, dock_area, ev: &DockEvent, window, cx| match ev {
                DockEvent::LayoutChanged => this.save_layout(dock_area, window, cx),
            },
        )
        .detach();

        cx.on_app_quit({
            let dock_area = dock_area.clone();
            move |_, cx| {
                let state = dock_area.read(cx).dump(cx);
                cx.background_executor().spawn(async move {
                    // Save layout before quitting
                    Self::save_tiles(&state).unwrap();
                })
            }
        })
        .detach();

        Self {
            dock_area,
            last_layout_state: None,
            _save_layout_task: None,
        }
    }

    fn save_layout(
        &mut self,
        dock_area: &Entity<DockArea>,
        _: &mut Window,
        cx: &mut Context<Self>,
    ) {
        let dock_area = dock_area.clone();
        self._save_layout_task = Some(cx.spawn(|this, cx| async move {
            Timer::after(Duration::from_secs(10)).await;

            let _ = cx.update(|cx| {
                let dock_area = dock_area.read(cx);
                let state = dock_area.dump(cx);

                let last_layout_state = this.upgrade().unwrap().read(cx).last_layout_state.clone();
                if Some(&state) == last_layout_state.as_ref() {
                    return;
                }

                Self::save_tiles(&state).unwrap();
                let _ = this.update(cx, |this, _| {
                    this.last_layout_state = Some(state);
                });
            });
        }));
    }

    fn save_tiles(state: &DockAreaState) -> Result<()> {
        println!("Save tiles...");
        let json = serde_json::to_string_pretty(state)?;
        std::fs::write("target/tiles.json", json)?;
        Ok(())
    }

    fn load_tiles(
        dock_area: Entity<DockArea>,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) -> Result<()> {
        let fname = "target/tiles.json";
        let json = std::fs::read_to_string(fname)?;
        let state = serde_json::from_str::<DockAreaState>(&json)?;

        // Check if the saved layout version is different from the current version
        // Notify the user and ask if they want to reset the layout to default.
        if state.version != Some(TILES_DOCK_AREA.version) {
            let answer = window.prompt(
                PromptLevel::Info,
                "The default tiles layout has been updated.\n\
                Do you want to reset the layout to default?",
                None,
                &["Yes", "No"],
                cx,
            );

            let weak_dock_area = dock_area.downgrade();
            cx.spawn_in(window, |this, mut window| async move {
                if answer.await == Ok(0) {
                    _ = this.update_in(&mut window, |_, window, cx| {
                        Self::reset_default_layout(weak_dock_area, window, cx);
                    });
                }
            })
            .detach();
        }

        dock_area.update(cx, |dock_area, cx| {
            dock_area.load(state, window, cx).context("load layout")?;

            Ok::<(), anyhow::Error>(())
        })
    }

    fn reset_default_layout(
        dock_area: WeakEntity<DockArea>,
        window: &mut Window,
        cx: &mut Context<Self>,
    ) {
        let dock_item = Self::init_default_layout(&dock_area, window, cx);
        _ = dock_area.update(cx, |view, cx| {
            view.set_version(TILES_DOCK_AREA.version, window, cx);
            view.set_center(dock_item, window, cx);

            Self::save_tiles(&view.dump(cx)).unwrap();
        });
    }

    fn init_default_layout(
        dock_area: &WeakEntity<DockArea>,
        window: &mut Window,
        cx: &mut App,
    ) -> DockItem {
        DockItem::tiles(
            vec![
                DockItem::tab(
                    StoryContainer::panel::<ButtonStory>(window, cx),
                    dock_area,
                    window,
                    cx,
                ),
                DockItem::tab(
                    StoryContainer::panel::<IconStory>(window, cx),
                    dock_area,
                    window,
                    cx,
                ),
            ],
            vec![
                Bounds::new(point(px(10.), px(10.)), size(px(610.), px(190.))),
                Bounds::new(point(px(120.), px(10.)), size(px(650.), px(300.))),
            ],
            dock_area,
            window,
            cx,
        )
    }

    pub fn new_local(cx: &mut App) -> Task<anyhow::Result<WindowHandle<Root>>> {
        let mut window_size = size(px(1600.0), px(1200.0));
        if let Some(display) = cx.primary_display() {
            let display_size = display.bounds().size;
            window_size.width = window_size.width.min(display_size.width * 0.85);
            window_size.height = window_size.height.min(display_size.height * 0.85);
        }
        let window_bounds = Bounds::centered(None, window_size, cx);

        cx.spawn(|mut cx| async move {
            let options = WindowOptions {
                window_bounds: Some(WindowBounds::Windowed(window_bounds)),
                titlebar: Some(TitlebarOptions {
                    title: None,
                    appears_transparent: true,
                    traffic_light_position: Some(point(px(9.0), px(9.0))),
                }),
                window_min_size: Some(gpui::Size {
                    width: px(640.),
                    height: px(480.),
                }),
                kind: WindowKind::Normal,
                #[cfg(target_os = "linux")]
                window_background: gpui::WindowBackgroundAppearance::Transparent,
                #[cfg(target_os = "linux")]
                window_decorations: Some(gpui::WindowDecorations::Client),
                ..Default::default()
            };

            let window = cx.open_window(options, |window, cx| {
                let tiles_view = cx.new(|cx| Self::new(window, cx));
                cx.new(|cx| Root::new(tiles_view.into(), window, cx))
            })?;

            window
                .update(&mut cx, |_, window, _| {
                    window.activate_window();
                    window.set_window_title("Story Tiles");
                })
                .expect("failed to update window");

            Ok(window)
        })
    }
}

pub fn open_new(
    cx: &mut App,
    init: impl FnOnce(&mut Root, &mut Window, &mut Context<Root>) + 'static + Send,
) -> Task<()> {
    let task: Task<std::result::Result<WindowHandle<Root>, anyhow::Error>> =
        StoryTiles::new_local(cx);
    cx.spawn(|mut cx| async move {
        if let Some(root) = task.await.ok() {
            root.update(&mut cx, |workspace, window, cx| init(workspace, window, cx))
                .expect("failed to init workspace");
        }
    })
}

impl Render for StoryTiles {
    fn render(&mut self, window: &mut Window, cx: &mut Context<Self>) -> impl IntoElement {
        let drawer_layer = Root::render_drawer_layer(window, cx);
        let modal_layer = Root::render_modal_layer(window, cx);
        let notification_layer = Root::render_notification_layer(window, cx);

        div()
            .font_family(".SystemUIFont")
            .relative()
            .size_full()
            .flex()
            .flex_col()
            .bg(cx.theme().background)
            .text_color(cx.theme().foreground)
            .child(TitleBar::new().child(div().flex().items_center().child("Story Tiles")))
            .child(self.dock_area.clone())
            .children(drawer_layer)
            .children(modal_layer)
            .child(div().absolute().top_8().children(notification_layer))
    }
}

fn main() {
    let app = Application::new().with_assets(Assets);

    app.run(move |cx| {
        ui::init(cx);
        story::init(cx);

        cx.on_action(quit);

        cx.set_menus(vec![Menu {
            name: "GPUI App".into(),
            items: vec![MenuItem::action("Quit", Quit)],
        }]);
        cx.activate(true);

        open_new(cx, |_, _, _| {
            // do something
        })
        .detach();
    });
}

fn quit(_: &Quit, cx: &mut App) {
    cx.quit();
}
