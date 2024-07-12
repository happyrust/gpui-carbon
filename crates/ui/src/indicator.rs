use std::time::Duration;

use crate::{theme::ActiveTheme, Icon, IconName, Size};
use gpui::{
    div, ease_in_out, percentage, Animation, AnimationExt as _, Hsla, IntoElement, ParentElement,
    RenderOnce, Styled as _, Transformation,
};

#[derive(IntoElement)]
pub struct Indicator {
    size: Size,
    icon: IconName,
    speed: Duration,
    color: Option<Hsla>,
}

impl Indicator {
    pub fn new() -> Self {
        Self {
            size: Size::Medium,
            speed: Duration::from_secs_f64(0.8),
            icon: IconName::Loader,
            color: None,
        }
    }

    pub fn size(mut self, size: impl Into<Size>) -> Self {
        self.size = size.into();
        self
    }

    pub fn icon(mut self, icon: IconName) -> Self {
        self.icon = icon;
        self
    }

    pub fn color(mut self, color: Hsla) -> Self {
        self.color = Some(color);
        self
    }
}

impl RenderOnce for Indicator {
    fn render(self, cx: &mut gpui::WindowContext) -> impl IntoElement {
        let color = self.color.unwrap_or_else(|| cx.theme().indicator);
        div()
            .child(
                Icon::new(self.icon)
                    .size(self.size)
                    .text_color(color)
                    .with_animation(
                        "circle",
                        Animation::new(self.speed).repeat().with_easing(ease_in_out),
                        |this, delta| this.transform(Transformation::rotate(percentage(delta))),
                    ),
            )
            .into_element()
    }
}
