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
    // To support Chinese characters on Windows, you need to set a default font
    // that contains Chinese glyphs. For iced 0.12, you would do this by
    // modifying the `Settings` struct.
    //
    // For example:
    //
    // let mut settings = Settings::default();
    // settings.default_font = Some(include_bytes!("../font.ttf"));
    // DocxApp::run(settings)
    //
    // Please ensure you have a font file named "font.ttf" in the project root
    let mut settings = Settings::default();
    settings.fonts.push(std::borrow::Cow::from(include_bytes!("../font.ttf")));
    settings.default_font = Font::with_name("Noto Sans SC");
    DocxApp::run(settings)
}

// --- Application State ---
struct DocxApp {
    file_path: Option<PathBuf>,
    created_date: String,
    modified_date: String,
    status_message: String,
    is_loading: bool,
}

// --- Messages to update state ---
#[derive(Debug, Clone)]
enum Message {
    SelectFile,
    FileSelected(Option<PathBuf>),
    FileLoaded(Result<(String, String), String>),
    CreatedDateChanged(String),
    ModifiedDateChanged(String),
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
                status_message: "Please select a .docx file to begin.".to_string(),
                is_loading: false,
            },
            Command::none(),
        )
    }

    fn title(&self) -> String {
        String::from("DOCX Metadata Editor")
    }

    fn update(&mut self, message: Message) -> Command<Message> {
        match message {
            Message::SelectFile => {
                self.is_loading = true;
                self.status_message = "Opening file dialog...".to_string();
                Command::perform(select_file_async(), Message::FileSelected)
            }
            Message::FileSelected(Some(path)) => {
                self.is_loading = true;
                self.status_message = format!("Loading metadata from {}...", path.display());
                self.file_path = Some(path.clone());
                Command::perform(load_metadata(path), Message::FileLoaded)
            }
            Message::FileSelected(None) => {
                self.is_loading = false;
                self.status_message = "File selection cancelled.".to_string();
                Command::none()
            }
            Message::FileLoaded(Ok((created, modified))) => {
                self.is_loading = false;
                self.created_date = created;
                self.modified_date = modified;
                self.status_message = "File loaded successfully.".to_string();
                Command::none()
            }
            Message::FileLoaded(Err(e)) => {
                self.is_loading = false;
                self.file_path = None;
                self.created_date.clear();
                self.modified_date.clear();
                self.status_message = format!("Error: {}", e);
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
            Message::SaveChanges => {
                if let Some(path) = self.file_path.clone() {
                    self.is_loading = true;
                    self.status_message = "Saving changes...".to_string();
                    let created = self.created_date.clone();
                    let modified = self.modified_date.clone();
                    Command::perform(save_metadata(path, created, modified), Message::FileSaved)
                } else {
                    self.status_message = "No file selected to save.".to_string();
                    Command::none()
                }
            }
            Message::FileSaved(Ok(())) => {
                self.is_loading = false;
                self.status_message = "File saved successfully!".to_string();
                Command::none()
            }
            Message::FileSaved(Err(e)) => {
                self.is_loading = false;
                self.status_message = format!("Error saving file: {}", e);
                Command::none()
            }
        }
    }

    fn view(&self) -> Element<'_, Message> {
        let file_display = self
            .file_path
            .as_ref()
            .map_or("No file selected", |p| p.to_str().unwrap_or("Invalid path"));

        let select_button = button("Select .docx File").on_press(Message::SelectFile);

        let mut save_button = button("Save Changes");
        if self.file_path.is_some() {
            save_button = save_button.on_press(Message::SaveChanges);
        }

        let content = column(vec![
            select_button.into(),
            text(file_display).size(16).into(),
            row(vec![
                text("Created Date:").width(Length::Fixed(120.0)).into(),
                text_input("e.g., 2023-01-01T12:00:00Z", &self.created_date)
                    .on_input(Message::CreatedDateChanged)
                    .into(),
            ])
            .spacing(10)
            .into(),
            row(vec![
                text("Modified Date:").width(Length::Fixed(120.0)).into(),
                text_input("e.g., 2023-01-01T13:00:00Z", &self.modified_date)
                    .on_input(Message::ModifiedDateChanged)
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
        .add_filter("Word Document", &["docx"])
        .pick_file()
}

async fn load_metadata(path: PathBuf) -> Result<(String, String), String> {
    let file = File::open(&path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    let mut core_props_entry = archive
        .by_name("docProps/core.xml")
        .map_err(|_| "docProps/core.xml not found in archive.".to_string())?;

    let mut core_props_buffer = Vec::new();
    core_props_entry
        .read_to_end(&mut core_props_buffer)
        .map_err(|e| e.to_string())?;
    let mut reader = Reader::from_reader(&core_props_buffer[..]);

    let mut created = String::new();
    let mut modified = String::new();
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
                _ => (),
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(format!("XML parsing error: {}", e)),
            _ => (),
        }
        buf.clear();
    }

    if created.is_empty() || modified.is_empty() {
        return Err("Could not find created/modified date tags in core.xml.".to_string());
    }

    Ok((created, modified))
}

async fn save_metadata(
    path: PathBuf,
    created_date: String,
    modified_date: String,
) -> Result<(), String> {
    // Validate date formats before proceeding
    DateTime::parse_from_rfc3339(&created_date.replace("Z", "+00:00")).map_err(|_| {
        "Invalid 'Created Date' format. Use ISO 8601 (e.g., YYYY-MM-DDTHH:MM:SSZ).".to_string()
    })?;
    DateTime::parse_from_rfc3339(&modified_date.replace("Z", "+00:00")).map_err(|_| {
        "Invalid 'Modified Date' format. Use ISO 8601 (e.g., YYYY-MM-DDTHH:MM:SSZ).".to_string()
    })?;

    let temp_path = path.with_extension("tmp");

    {
        // Scoped to ensure file handles are closed before rename
        let file = File::open(&path).map_err(|e| e.to_string())?;
        let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;

        let temp_file = File::create(&temp_path).map_err(|e| e.to_string())?;
        let mut zip_writer = ZipWriter::new(temp_file);

        let options: zip::write::FileOptions<'_, ()> =
            FileOptions::default().compression_method(zip::CompressionMethod::Deflated);

        for i in 0..archive.len() {
            let mut file = archive.by_index(i).unwrap();
            if file.name() == "docProps/core.xml" {
                // Skip the old core.xml, we will write a new one
                continue;
            }
            zip_writer
                .start_file(file.name(), options)
                .map_err(|e| e.to_string())?;
            let mut buffer = Vec::new();
            file.read_to_end(&mut buffer).map_err(|e| e.to_string())?;
            zip_writer.write_all(&buffer).map_err(|e| e.to_string())?;
        }

        // Now, create the modified core.xml
        let new_core_xml = generate_core_xml(&path, &created_date, &modified_date)?;
        zip_writer
            .start_file("docProps/core.xml", options)
            .map_err(|e| e.to_string())?;
        zip_writer
            .write_all(new_core_xml.as_bytes())
            .map_err(|e| e.to_string())?;

        zip_writer.finish().map_err(|e| e.to_string())?;
    }

    // Replace original file with the temp file
    fs::rename(&temp_path, &path).map_err(|e| format!("Failed to replace original file: {}", e))
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
        .map_err(|_| "docProps/core.xml not found.".to_string())?;

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
                let elem = e.to_owned();
                if e.name().as_ref() == b"dcterms:created"
                    || e.name().as_ref() == b"dcterms:modified"
                {
                    // Keep attributes, but prepare to write new text
                    writer.write_event(Event::Start(elem)).unwrap();
                    let text_to_write = if e.name().as_ref() == b"dcterms:created" {
                        new_created
                    } else {
                        new_modified
                    };
                    writer
                        .write_event(Event::Text(BytesText::new(text_to_write)))
                        .unwrap();
                    // We must skip the original text event, which comes next
                    reader.read_to_end_into(e.name(), &mut Vec::new()).unwrap();
                    writer.write_event(Event::End(e.to_end())).unwrap();
                } else {
                    writer.write_event(Event::Start(elem)).unwrap();
                }
            }
            Ok(Event::End(e)) => {
                // This case is handled by the logic above to ensure proper closing
                if e.name().as_ref() != b"dcterms:created"
                    && e.name().as_ref() != b"dcterms:modified"
                {
                    writer.write_event(Event::End(e.to_owned())).unwrap();
                }
            }
            Ok(Event::Eof) => break,
            Ok(e) => {
                writer.write_event(e).unwrap();
            }
            Err(e) => return Err(format!("XML processing error: {}", e)),
        }
        buf.clear();
    }

    let result = writer.into_inner().into_inner();
    String::from_utf8(result).map_err(|e| e.to_string())
}
