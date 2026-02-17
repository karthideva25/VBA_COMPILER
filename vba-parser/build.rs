use cc;
fn main() {
    println!("cargo:rerun-if-changed=src/parser.c");
    println!("cargo:rerun-if-changed=src/scanner.c");

   
    let mut build = cc::Build::new();
    build.include("src");
    build.file("src/parser.c");

    // 👇 Add this line to compile scanner.c
    if std::path::Path::new("src/scanner.c").exists() {
        build.file("src/scanner.c");
    }

    build.compile("tree_sitter_vba");
    // If you have scanner.c or scanner.cc, include this too:
    // .file("src/scanner.c") OR .file("src/scanner.cc")

    // println!("cargo:rustc-link-lib=static=tree-sitter-vba");
    // println!("cargo:rustc-link-search=native={}", std::env::var("OUT_DIR").unwrap());

}
