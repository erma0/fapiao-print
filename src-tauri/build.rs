fn main() {
    #[cfg(target_os = "windows")]
    {
        cc::Build::new()
            .file("src/seh_wrapper.c")
            .compile("seh_wrapper");
    }
    tauri_build::build()
}
