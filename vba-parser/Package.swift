// swift-tools-version:5.3

import Foundation
import PackageDescription

var sources = ["src/parser.c"]
if FileManager.default.fileExists(atPath: "src/scanner.c") {
    sources.append("src/scanner.c")
}

let package = Package(
    name: "TreeSitterVbatestjune",
    products: [
        .library(name: "TreeSitterVbatestjune", targets: ["TreeSitterVbatestjune"]),
    ],
    dependencies: [
        .package(url: "https://github.com/tree-sitter/swift-tree-sitter", from: "0.8.0"),
    ],
    targets: [
        .target(
            name: "TreeSitterVbatestjune",
            dependencies: [],
            path: ".",
            sources: sources,
            resources: [
                .copy("queries")
            ],
            publicHeadersPath: "bindings/swift",
            cSettings: [.headerSearchPath("src")]
        ),
        .testTarget(
            name: "TreeSitterVbatestjuneTests",
            dependencies: [
                "SwiftTreeSitter",
                "TreeSitterVbatestjune",
            ],
            path: "bindings/swift/TreeSitterVbatestjuneTests"
        )
    ],
    cLanguageStandard: .c11
)
