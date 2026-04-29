use std::fs;

fn main() {
    // 测试高德打车发票OCR
    let pdf1 = r"C:\Users\Administrator\Downloads\【K9用车-177.86元-1个行程】高德打车电子发票.pdf";
    println!("=== 高德打车发票 ===");
    println!("Path exists: {}", fs::metadata(pdf1).is_ok());

    // 测试火车票OCR
    let pdf2 = r"C:\Users\Administrator\Downloads\25329116804007140998.pdf";
    println!("\n=== 火车票 ===");
    println!("Path exists: {}", fs::metadata(pdf2).is_ok());
}
