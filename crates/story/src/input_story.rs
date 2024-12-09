use gpui::{
    actions, div, px, AppContext, FocusHandle, InteractiveElement, IntoElement, KeyBinding,
    ParentElement as _, Render, SharedString, Styled, View, ViewContext, VisualContext,
    WindowContext,
};
use regex::Regex;

use crate::section;
use ui::{
    button::{Button, ButtonVariant, ButtonVariants as _},
    checkbox::Checkbox,
    h_flex,
    input::{InputEvent, OtpInput, TextInput},
    number_input::{NumberInput, NumberInputEvent},
    prelude::FluentBuilder as _,
    scroll::ScrollbarAxis,
    v_flex, FocusableCycle, IconName, Sizable, StyledExt,
};

actions!(input_story, [Tab, TabPrev]);

const CONTEXT: &str = "InputStory";

pub fn init(cx: &mut AppContext) {
    cx.bind_keys([
        KeyBinding::new("shift-tab", TabPrev, Some(CONTEXT)),
        KeyBinding::new("tab", Tab, Some(CONTEXT)),
    ])
}

pub struct InputStory {
    input1: View<TextInput>,
    input2: View<TextInput>,
    textarea: View<TextInput>,
    number_input1_value: i64,
    number_input1: View<NumberInput>,
    number_input2: View<NumberInput>,
    number_input2_value: u64,
    mash_input: View<TextInput>,
    disabled_input: View<TextInput>,
    prefix_input1: View<TextInput>,
    suffix_input1: View<TextInput>,
    both_input1: View<TextInput>,
    large_input: View<TextInput>,
    small_input: View<TextInput>,
    otp_masked: bool,
    otp_input: View<OtpInput>,
    otp_value: Option<SharedString>,
    otp_input_small: View<OtpInput>,
    otp_input_large: View<OtpInput>,
    opt_input_sized: View<OtpInput>,
}

impl super::Story for InputStory {
    fn title() -> &'static str {
        "Input"
    }

    fn closeable() -> bool {
        false
    }

    fn new_view(cx: &mut WindowContext) -> View<impl gpui::FocusableView> {
        Self::view(cx)
    }
}

impl InputStory {
    pub fn view(cx: &mut WindowContext) -> View<Self> {
        cx.new_view(Self::new)
    }

    fn new(cx: &mut ViewContext<Self>) -> Self {
        let input1 = cx.new_view(|cx| {
            let mut input = TextInput::new(cx).cleanable();
            input.set_text(
                "Hello 世界，this is GPUI component, this is a long text.",
                cx,
            );
            input
        });
        cx.subscribe(&input1, Self::on_input_event).detach();

        let input2 = cx.new_view(|cx| TextInput::new(cx).placeholder("Enter text here..."));
        cx.subscribe(&input2, Self::on_input_event).detach();

        let textarea = cx.new_view(|cx| {
            let mut input = TextInput::new(cx)
                .multi_line()
                .rows(10)
                .placeholder("Enter text here...");
            input.set_text(
                unindent::unindent(
                    r#"Hello 世界，this is GPUI component.

                The GPUI Component is a collection of UI components for GPUI framework, including.

                Button, Input, Checkbox, Radio, Dropdown, Tab, and more...

                Here is an application that is built by using GPUI Component.

                > This application is still under development, not published yet.

                ![image](https://github.com/user-attachments/assets/559a648d-19df-4b5a-b563-b78cc79c8894)

                ![image](https://github.com/user-attachments/assets/5e06ad5d-7ea0-43db-8d13-86a240da4c8d)

                ## Demo

                If you want to see the demo, here is a some demo applications.
                "#,
                ),
                cx,
            );
            input
        });
        cx.subscribe(&textarea, Self::on_input_event).detach();

        let number_input1_value = 1;
        let number_input1 = cx.new_view(|cx| {
            let input = NumberInput::new(cx).placeholder("Number Input", cx);
            input.set_value(number_input1_value.to_string(), cx);
            input
        });
        cx.subscribe(&number_input1, Self::on_number_input1_event)
            .detach();

        let number_input2 = cx.new_view(|cx| {
            NumberInput::new(cx)
                .placeholder("Unsized Integer Number Input", cx)
                .pattern(Regex::new(r"^\d+$").unwrap(), cx)
                .small(cx)
        });

        cx.subscribe(&number_input2, Self::on_number_input2_event)
            .detach();

        let mask_input = cx.new_view(|cx| {
            let mut input = TextInput::new(cx).cleanable();
            input.set_masked(true, cx);
            input.set_text("this-is-password", cx);
            input
        });

        let prefix_input1 = cx.new_view(|cx| {
            TextInput::new(cx)
                .prefix(|_| div().child(IconName::Search).ml_3())
                .placeholder("Search some thing...")
                .cleanable()
        });
        let suffix_input1 = cx.new_view(|cx| {
            TextInput::new(cx)
                .suffix(|_| div().child(IconName::Info).mr_3())
                .placeholder("This input only support [a-zA-Z0-9] characters.")
                .pattern(regex::Regex::new(r"^[a-zA-Z0-9]*$").unwrap())
                .cleanable()
        });
        let both_input1 = cx.new_view(|cx| {
            TextInput::new(cx)
                .prefix(|_| div().child(IconName::Search).ml_3())
                .suffix(|_| div().child(IconName::Info).mr_3())
                .cleanable()
                .placeholder("This input have prefix and suffix.")
        });

        let otp_input = cx.new_view(|cx| OtpInput::new(6, cx).masked(true));
        cx.subscribe(&otp_input, |this, _, ev: &InputEvent, cx| match ev {
            InputEvent::Change(text) => {
                this.otp_value = Some(text.clone());
                cx.notify();
            }
            _ => {}
        })
        .detach();

        Self {
            input1,
            input2,
            textarea,
            number_input1,
            number_input1_value,
            number_input2,
            number_input2_value: 0,
            mash_input: mask_input,
            disabled_input: cx.new_view(|cx| {
                let mut input = TextInput::new(cx);
                input.set_text("This is disabled input", cx);
                input.set_disabled(true, cx);
                input
            }),
            large_input: cx.new_view(|cx| TextInput::new(cx).large().placeholder("Large input")),
            small_input: cx.new_view(|cx| {
                TextInput::new(cx)
                    .small()
                    .validate(|s| s.parse::<f32>().is_ok())
                    .placeholder("validate to limit float number.")
            }),
            prefix_input1,
            suffix_input1,
            both_input1,
            otp_masked: true,
            otp_input,
            otp_value: None,
            otp_input_small: cx.new_view(|cx| {
                OtpInput::new(6, cx)
                    .default_value("123456")
                    .masked(true)
                    .small()
                    .groups(1)
            }),
            otp_input_large: cx.new_view(|cx| {
                OtpInput::new(6, cx)
                    .groups(3)
                    .large()
                    .default_value("012345")
                    .masked(true)
            }),
            opt_input_sized: cx.new_view(|cx| {
                OtpInput::new(4, cx)
                    .groups(1)
                    .masked(true)
                    .default_value("654321")
                    .with_size(px(55.))
            }),
        }
    }

    fn tab(&mut self, _: &Tab, cx: &mut ViewContext<Self>) {
        self.cycle_focus(true, cx);
    }

    fn tab_prev(&mut self, _: &TabPrev, cx: &mut ViewContext<Self>) {
        self.cycle_focus(false, cx);
    }

    fn on_input_event(
        &mut self,
        _: View<TextInput>,
        event: &InputEvent,
        _cx: &mut ViewContext<Self>,
    ) {
        match event {
            InputEvent::Change(text) => println!("Change: {}", text),
            InputEvent::PressEnter => println!("PressEnter"),
            InputEvent::Focus => println!("Focus"),
            InputEvent::Blur => println!("Blur"),
        };
    }

    fn on_number_input1_event(
        &mut self,
        _: View<NumberInput>,
        event: &NumberInputEvent,
        cx: &mut ViewContext<Self>,
    ) {
        match event {
            NumberInputEvent::Input(input_event) => match input_event {
                InputEvent::Change(text) => println!("Change: {}", text),
                InputEvent::PressEnter => println!("PressEnter"),
                InputEvent::Focus => println!("Focus"),
                InputEvent::Blur => println!("Blur"),
            },
            NumberInputEvent::Step(step_action) => match step_action {
                ui::number_input::StepAction::Decrement => {
                    self.number_input1_value = self.number_input1_value - 1;
                    self.number_input1.update(cx, |input, cx| {
                        input.set_value(self.number_input1_value.to_string(), cx);
                    });
                }
                ui::number_input::StepAction::Increment => {
                    self.number_input1_value = self.number_input1_value + 1;
                    self.number_input1.update(cx, |input, cx| {
                        input.set_value(self.number_input1_value.to_string(), cx);
                    });
                }
            },
        }
    }

    fn on_number_input2_event(
        &mut self,
        _: View<NumberInput>,
        event: &NumberInputEvent,
        cx: &mut ViewContext<Self>,
    ) {
        match event {
            NumberInputEvent::Input(input_event) => match input_event {
                InputEvent::Change(text) => println!("Change: {}", text),
                InputEvent::PressEnter => println!("PressEnter"),
                InputEvent::Focus => println!("Focus"),
                InputEvent::Blur => println!("Blur"),
            },
            NumberInputEvent::Step(step_action) => match step_action {
                ui::number_input::StepAction::Decrement => {
                    if self.number_input2_value.le(&0) {
                        return;
                    }

                    self.number_input2_value = self.number_input2_value - 1;
                    self.number_input2.update(cx, |input, cx| {
                        input.set_value(self.number_input2_value.to_string(), cx);
                    });
                }
                ui::number_input::StepAction::Increment => {
                    self.number_input2_value = self.number_input2_value + 1;
                    self.number_input2.update(cx, |input, cx| {
                        input.set_value(self.number_input2_value.to_string(), cx);
                    });
                }
            },
        }
    }

    fn toggle_opt_masked(&mut self, _: &bool, cx: &mut ViewContext<Self>) {
        self.otp_masked = !self.otp_masked;
        self.otp_input
            .update(cx, |input, cx| input.set_masked(self.otp_masked, cx));
        self.otp_input_small
            .update(cx, |input, cx| input.set_masked(self.otp_masked, cx));
        self.otp_input_large
            .update(cx, |input, cx| input.set_masked(self.otp_masked, cx));
        self.opt_input_sized
            .update(cx, |input, cx| input.set_masked(self.otp_masked, cx));
    }
}

impl FocusableCycle for InputStory {
    fn cycle_focus_handles(&self, cx: &mut ViewContext<Self>) -> Vec<FocusHandle> {
        [
            self.input1.focus_handle(cx),
            self.input2.focus_handle(cx),
            self.disabled_input.focus_handle(cx),
            self.mash_input.focus_handle(cx),
            self.prefix_input1.focus_handle(cx),
            self.both_input1.focus_handle(cx),
            self.suffix_input1.focus_handle(cx),
            self.large_input.focus_handle(cx),
            self.small_input.focus_handle(cx),
            self.otp_input.focus_handle(cx),
        ]
        .to_vec()
    }
}
impl gpui::FocusableView for InputStory {
    fn focus_handle(&self, cx: &gpui::AppContext) -> gpui::FocusHandle {
        self.input1.focus_handle(cx)
    }
}

impl Render for InputStory {
    fn render(&mut self, cx: &mut ViewContext<Self>) -> impl IntoElement {
        v_flex()
            .key_context(CONTEXT)
            .id("input-story")
            .scrollable(cx.entity_id(), ScrollbarAxis::Vertical)
            .on_action(cx.listener(Self::tab))
            .on_action(cx.listener(Self::tab_prev))
            .size_full()
            .p_4()
            .justify_start()
            .gap_3()
            .child(
                h_flex()
                    .gap_3()
                    .items_start()
                    .child(
                        section("Normal Input", cx)
                            .child(self.input1.clone())
                            .child(self.input2.clone())
                            .child(
                                v_flex()
                                    .gap_y_4()
                                    .w_full()
                                    .child("Number Input")
                                    .child(self.number_input1.clone())
                                    .child(self.number_input2.clone()),
                            ),
                    )
                    .child(section("Textarea", cx).child(self.textarea.clone()))
                    .child(
                        section("Input State", cx)
                            .child(self.disabled_input.clone())
                            .child(self.mash_input.clone()),
                    ),
            )
            .child(
                h_flex()
                    .gap_3()
                    .items_start()
                    .child(
                        section("Prefix and Suffix", cx)
                            .child(self.prefix_input1.clone())
                            .child(self.both_input1.clone())
                            .child(self.suffix_input1.clone()),
                    )
                    .child(
                        section("Input Size", cx)
                            .child(self.large_input.clone())
                            .child(self.small_input.clone()),
                    ),
            )
            .child(
                section(
                    h_flex()
                        .items_center()
                        .justify_between()
                        .child("OTP Input")
                        .child(
                            Checkbox::new("otp-mask")
                                .label("Masked")
                                .checked(self.otp_masked)
                                .on_click(cx.listener(Self::toggle_opt_masked)),
                        ),
                    cx,
                )
                .child(
                    v_flex()
                        .gap_3()
                        .child(self.otp_input_small.clone())
                        .child(self.otp_input.clone())
                        .when_some(self.otp_value.clone(), |this, otp| {
                            this.child(format!("Your OTP: {}", otp))
                        })
                        .child(self.otp_input_large.clone())
                        .child(self.opt_input_sized.clone()),
                ),
            )
            .child(
                h_flex()
                    .items_center()
                    .w_full()
                    .gap_3()
                    .child(
                        Button::new("btn-submit")
                            .w_full()
                            .with_variant(ButtonVariant::Primary)
                            .label("Submit")
                            .on_click(cx.listener(|_, _, cx| cx.dispatch_action(Box::new(Tab)))),
                    )
                    .child(
                        Button::new("btn-cancel")
                            .w_full()
                            .label("Cancel")
                            .into_element(),
                    ),
            )
    }
}
