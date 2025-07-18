use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use chrono::DateTime;
use iced::{
    executor,
    widget::{button, column, container, row, text, text_input},
    Application, Command, Element, Font, Length, Settings, Theme,
};
use quick_xml::events::{BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use rfd::FileDialog;
use zip::write::{FileOptions, ZipWriter};
use zip::ZipArchive;

// --- Main Application Entry Point ---
pub fn main() -> iced::Result {
    let mut settings = Settings::default();
    settings
        .fonts
        .push(std::borrow::Cow::from(include_bytes!("../font.ttf")));
    settings.default_font = Font::with_name("Noto Sans SC");
    DocxApp::run(settings)
}

// --- Application State ---
struct DocxApp {
    file_path: Option<PathBuf>,
    created_date: String,
    modified_date: String,
    last_printed_date: String, // New field for last printed date
    status_message: String,
    is_loading: bool,
}

// --- Messages to update state ---
#[derive(Debug, Clone)]
enum Message {
    SelectFile,
    FileSelected(Option<PathBuf>),
    FileLoaded(Result<(String, String, String), String>), // Updated to include last printed date
    CreatedDateChanged(String),
    ModifiedDateChanged(String),
    LastPrintedDateChanged(String), // New message for last printed date
    SaveChanges,
    FileSaved(Result<(), String>),
}

// --- Iced Application Implementation ---
impl Application for DocxApp {
    type Executor = executor::Default;
    type Message = Message;
    type Theme = Theme;
    type Flags = ();

    fn new(_flags: ()) -> (Self, Command<Message>) {
        (
            Self {
                file_path: None,
                created_date: String::new(),
                modified_date: String::new(),
                last_printed_date: String::new(), // Initialize new field
                status_message: "请选择一个 .docx 文件开始".to_string(),
                is_loading: false,
            },
            Command::none(),
        )
    }

    fn title(&self) -> String {
        String::from("DOCX 元数据编辑器")
    }

    fn update(&mut self, message: Message) -> Command<Message> {
        match message {
            Message::SelectFile => {
                self.is_loading = true;
                self.status_message = "正在打开文件对话框...".to_string();
                Command::perform(select_file_async(), Message::FileSelected)
            }
            Message::FileSelected(Some(path)) => {
                self.is_loading = true;
                self.status_message = format!("正在从 {} 加载元数据...", path.display());
                self.file_path = Some(path.clone());
                Command::perform(load_metadata(path), Message::FileLoaded)
            }
            Message::FileSelected(None) => {
                self.is_loading = false;
                self.status_message = "文件选择已取消.".to_string();
                Command::none()
            }
            Message::FileLoaded(Ok((created, modified, last_printed))) => {
                self.is_loading = false;
                self.created_date = created;
                self.modified_date = modified;
                self.last_printed_date = last_printed; // Store last printed date
                self.status_message = "文件加载成功.".to_string();
                Command::none()
            }
            Message::FileLoaded(Err(e)) => {
                self.is_loading = false;
                self.file_path = None;
                self.created_date.clear();
                self.modified_date.clear();
                self.last_printed_date.clear(); // Clear last printed date on error
                self.status_message = format!("错误: {}", e);
                Command::none()
            }
            Message::CreatedDateChanged(date) => {
                self.created_date = date;
                Command::none()
            }
            Message::ModifiedDateChanged(date) => {
                self.modified_date = date;
                Command::none()
            }
            Message::LastPrintedDateChanged(date) => {
                self.last_printed_date = date; // Handle changes to last printed date
                Command::none()
            }
            Message::SaveChanges => {
                if let Some(path) = self.file_path.clone() {
                    self.is_loading = true;
                    self.status_message = "正在保存更改...".to_string();
                    let created = self.created_date.clone();
                    let modified = self.modified_date.clone();
                    let last_printed = self.last_printed_date.clone();
                    Command::perform(
                        save_metadata(path, created, modified, last_printed),
                        Message::FileSaved,
                    )
                } else {
                    self.status_message = "未选择要保存的文件.".to_string();
                    Command::none()
                }
            }
            Message::FileSaved(Ok(())) => {
                self.is_loading = false;
                self.status_message = "文件保存成功!".to_string();
                Command::none()
            }
            Message::FileSaved(Err(e)) => {
                self.is_loading = false;
                self.status_message = format!("保存文件时出错: {}", e);
                Command::none()
            }
        }
    }

    fn view(&self) -> Element<'_, Message> {
        let file_display = self
            .file_path
            .as_ref()
            .map_or("未选择文件", |p| p.to_str().unwrap_or("无效路径"));

        let select_button = button("选择 .docx 文件").on_press(Message::SelectFile);

        let mut save_button = button("保存更改");
        if self.file_path.is_some() {
            save_button = save_button.on_press(Message::SaveChanges);
        }

        let content = column(vec![
            select_button.into(),
            text(file_display).size(16).into(),
            row(vec![
                text("创建日期:").width(Length::Fixed(120.0)).into(),
                text_input("例如, 2023-01-01T12:00:00Z", &self.created_date)
                    .on_input(Message::CreatedDateChanged)
                    .into(),
            ])
            .spacing(10)
            .into(),
            row(vec![
                text("修改日期:").width(Length::Fixed(120.0)).into(),
                text_input("例如, 2023-01-01T13:00:00Z", &self.modified_date)
                    .on_input(Message::ModifiedDateChanged)
                    .into(),
            ])
            .spacing(10)
            .into(),
            row(vec![
                text("最后打印:").width(Length::Fixed(120.0)).into(),
                text_input("例如, 2023-01-01T14:00:00Z", &self.last_printed_date)
                    .on_input(Message::LastPrintedDateChanged)
                    .into(),
            ])
            .spacing(10)
            .into(),
            save_button.into(),
            text(&self.status_message).size(16).into(),
        ])
        .spacing(20)
        .padding(20);

        container(content)
            .width(Length::Fill)
            .height(Length::Fill)
            .center_x()
            .center_y()
            .into()
    }
}

// --- Async Helper Functions ---

async fn select_file_async() -> Option<PathBuf> {
    FileDialog::new()
        .add_filter("Word 文档", &["docx"])
        .pick_file()
}

async fn load_metadata(path: PathBuf) -> Result<(String, String, String), String> {
    let file = File::open(&path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;

    let (created, modified, last_printed) = {
        let mut core_props_entry = archive
            .by_name("docProps/core.xml")
            .map_err(|_| "在压缩包中找不到 docProps/core.xml。".to_string())?;
        let mut core_props_buffer = Vec::new();
        core_props_entry
            .read_to_end(&mut core_props_buffer)
            .map_err(|e| e.to_string())?;
        let mut reader = Reader::from_reader(&core_props_buffer[..]);
        let mut created = String::new();
        let mut modified = String::new();
        let mut last_printed = String::new();
        let mut buf = Vec::new();
        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => match e.name().as_ref() {
                    b"dcterms:created" => {
                        created = reader.read_text(e.name()).unwrap_or_default().to_string();
                    }
                    b"dcterms:modified" => {
                        modified = reader.read_text(e.name()).unwrap_or_default().to_string();
                    }
                    b"cp:lastPrinted" => {
                        last_printed = reader.read_text(e.name()).unwrap_or_default().to_string();
                    }
                    _ => (),
                },
                Ok(Event::Eof) => break,
                Err(e) => return Err(format!("core.xml XML 解析错误: {}", e)),
                _ => (),
            }
            buf.clear();
        }
        (created, modified, last_printed)
    };

    Ok((created, modified, last_printed))
}

async fn save_metadata(
    path: PathBuf,
    created_date: String,
    modified_date: String,
    last_printed_date: String,
) -> Result<(), String> {
    // Validate date formats before proceeding
    DateTime::parse_from_rfc3339(&created_date.replace("Z", "+00:00")).map_err(|_| {
        "创建日期' 格式无效。请使用 ISO 8601 (例如：YYYY-MM-DDTHH:MM:SSZ)。".to_string()
    })?;
    DateTime::parse_from_rfc3339(&modified_date.replace("Z", "+00:00")).map_err(|_| {
        "修改日期' 格式无效。请使用 ISO 8601 (例如：YYYY-MM-DDTHH:MM:SSZ)。".to_string()
    })?;
    if !last_printed_date.is_empty() {
        DateTime::parse_from_rfc3339(&last_printed_date.replace("Z", "+00:00")).map_err(|_| {
            "最后打印日期' 格式无效。请使用 ISO 8601 (例如：YYYY-MM-DDTHH:MM:SSZ)。".to_string()
        })?;
    }

    let temp_path = path.with_extension("tmp");

    {
        let file = File::open(&path).map_err(|e| e.to_string())?;
        let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
        let temp_file = File::create(&temp_path).map_err(|e| e.to_string())?;
        let mut zip_writer = ZipWriter::new(temp_file);
        let options: zip::write::FileOptions<'_, ()> =
            FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            let file_name = file.name();
            if file_name == "docProps/core.xml" || file_name == "docProps/app.xml" {
                continue; // Skip old property files
            }
            zip_writer
                .start_file(file.name(), options)
                .map_err(|e| e.to_string())?;
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer).map_err(|e| e.to_string())?;
            zip_writer.write_all(&buffer).map_err(|e| e.to_string())?;
        }

        // Create and write the modified core.xml
        let new_core_xml = generate_core_xml(&path, &created_date, &modified_date)?;
        zip_writer
            .start_file("docProps/core.xml", options)
            .map_err(|e| e.to_string())?;
        zip_writer
            .write_all(new_core_xml.as_bytes())
            .map_err(|e| e.to_string())?;

        // Create and write the modified app.xml
        let new_app_xml = generate_app_xml(&path, &last_printed_date)?;
        zip_writer
            .start_file("docProps/app.xml", options)
            .map_err(|e| e.to_string())?;
        zip_writer
            .write_all(new_app_xml.as_bytes())
            .map_err(|e| e.to_string())?;

        zip_writer.finish().map_err(|e| e.to_string())?;
    }

    fs::rename(&temp_path, &path).map_err(|e| format!("替换原始文件失败: {}", e))
}

fn generate_core_xml(
    original_path: &Path,
    new_created: &str,
    new_modified: &str,
) -> Result<String, String> {
    let file = File::open(original_path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    let mut core_props_entry = archive
        .by_name("docProps/core.xml")
        .map_err(|_| "找不到 docProps/core.xml。".to_string())?;

    let mut core_props_buffer = Vec::new();
    core_props_entry
        .read_to_end(&mut core_props_buffer)
        .map_err(|e| e.to_string())?;
    let mut reader = Reader::from_reader(&core_props_buffer[..]);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                let elem_name = e.name();
                let should_replace = elem_name.as_ref() == b"dcterms:created"
                    || elem_name.as_ref() == b"dcterms:modified";

                writer.write_event(Event::Start(e.to_owned())).unwrap();

                if should_replace {
                    let text_to_write = if elem_name.as_ref() == b"dcterms:created" {
                        new_created
                    } else {
                        new_modified
                    };
                    writer
                        .write_event(Event::Text(BytesText::new(text_to_write)))
                        .unwrap();
                    // Skip the original text event by reading until the end of the element
                    reader.read_to_end_into(elem_name, &mut Vec::new()).unwrap();
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                writer.write_event(e).unwrap();
            }
            Err(e) => return Err(format!("XML (core) 处理错误: {}", e)),
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner().into_inner()).map_err(|e| e.to_string())
}

fn generate_app_xml(original_path: &Path, new_last_printed: &str) -> Result<String, String> {
    let file = File::open(original_path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;

    // app.xml is optional, so we handle its absence gracefully.
    let app_props_buffer = match archive.by_name("docProps/app.xml") {
        Ok(mut entry) => {
            let mut buffer = Vec::new();
            entry.read_to_end(&mut buffer).map_err(|e| e.to_string())?;
            buffer
        }
        Err(_) => {
            // If app.xml doesn't exist, create a default structure.
            return Ok(format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <LastPrinted>{}</LastPrinted>
</Properties>"#,
                new_last_printed
            ));
        }
    };

    let mut reader = Reader::from_reader(&app_props_buffer[..]);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buf = Vec::new();
    let mut found_last_printed = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"LastPrinted" {
                    found_last_printed = true;
                    writer.write_event(Event::Start(e.to_owned())).unwrap();
                    writer
                        .write_event(Event::Text(BytesText::new(new_last_printed)))
                        .unwrap();
                    reader.read_to_end_into(e.name(), &mut Vec::new()).unwrap();
                } else {
                    writer.write_event(Event::Start(e.to_owned())).unwrap();
                }
            }
            Ok(Event::End(e)) => {
                // If we are at the end of the root and haven't found the tag, add it.
                if e.name().as_ref() == b"Properties"
                    && !found_last_printed
                    && !new_last_printed.is_empty()
                {
                    writer
                        .create_element("LastPrinted")
                        .write_text_content(BytesText::new(new_last_printed))
                        .unwrap();
                }
                writer.write_event(Event::End(e.to_owned())).unwrap();
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                writer.write_event(e).unwrap();
            }
            Err(e) => return Err(format!("XML (app) 处理错误: {}", e)),
        }
        buf.clear();
    }

    String::from_utf8(writer.into_inner().into_inner()).map_err(|e| e.to_string())
}
