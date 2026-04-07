//
//  ViewController.swift
//  Couturier
//

import Cocoa

// MARK: - Model

struct FontTheme {
    var name: String
    var headingFont: String
    var bodyFont: String
    var fileURL: URL?
}

// MARK: - XML Parser

private class FontThemeXMLParser: NSObject, XMLParserDelegate {
    private(set) var name = ""
    private(set) var headingFont = ""
    private(set) var bodyFont = ""
    private var inMajor = false
    private var inMinor = false

    func parse(data: Data) -> Bool {
        let parser = XMLParser(data: data)
        parser.delegate = self
        return parser.parse()
    }

    func parser(_ parser: XMLParser, didStartElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?,
                attributes attr: [String: String] = [:]) {
        switch elementName {
        case "a:fontScheme":
            name = attr["name"] ?? ""
        case "a:majorFont":
            inMajor = true; inMinor = false
        case "a:minorFont":
            inMinor = true; inMajor = false
        case "a:latin":
            if inMajor { headingFont = attr["typeface"] ?? "" }
            else if inMinor { bodyFont = attr["typeface"] ?? "" }
        default: break
        }
    }

    func parser(_ parser: XMLParser, didEndElement elementName: String,
                namespaceURI: String?, qualifiedName qName: String?) {
        if elementName == "a:majorFont" { inMajor = false }
        if elementName == "a:minorFont" { inMinor = false }
    }
}

// MARK: - Theme Manager

enum ThemeManagerError: LocalizedError {
    case folderCreationFailed
    case saveFailed(Error)

    var errorDescription: String? {
        switch self {
        case .folderCreationFailed: return "Could not create the Theme Fonts folder."
        case .saveFailed(let e): return "Save failed: \(e.localizedDescription)"
        }
    }
}

struct FontThemeManager {

    static var themeFolderURL: URL {
        FileManager.default.homeDirectoryForCurrentUser
            .appendingPathComponent("Library/Group Containers/UBF8T346G9.Office/User Content/Themes/Theme Fonts")
    }

    static func loadThemes() -> [FontTheme] {
        let fm = FileManager.default
        guard let urls = try? fm.contentsOfDirectory(at: themeFolderURL,
                                                      includingPropertiesForKeys: nil) else { return [] }
        return urls
            .filter { $0.pathExtension.lowercased() == "xml" }
            .compactMap { url -> FontTheme? in
                guard let data = try? Data(contentsOf: url) else { return nil }
                let p = FontThemeXMLParser()
                guard p.parse(data: data), !p.name.isEmpty else { return nil }
                return FontTheme(name: p.name, headingFont: p.headingFont,
                                 bodyFont: p.bodyFont, fileURL: url)
            }
            .sorted { $0.name.localizedCaseInsensitiveCompare($1.name) == .orderedAscending }
    }

    @discardableResult
    static func save(theme: FontTheme) throws -> URL {
        let fm = FileManager.default
        if !fm.fileExists(atPath: themeFolderURL.path) {
            try fm.createDirectory(at: themeFolderURL, withIntermediateDirectories: true)
        }
        let safeName = theme.name.isEmpty ? "Untitled" : theme.name
        let url = theme.fileURL ?? themeFolderURL.appendingPathComponent("\(safeName).xml")
        let xml = makeXML(name: safeName, headingFont: theme.headingFont, bodyFont: theme.bodyFont)
        try xml.write(to: url, atomically: true, encoding: .utf8)
        return url
    }

    static func rename(theme: inout FontTheme, newName: String) throws -> URL {
        let oldURL = theme.fileURL
        let newURL = themeFolderURL.appendingPathComponent("\(newName).xml")
        theme.name = newName
        theme.fileURL = newURL
        try save(theme: theme)
        if let old = oldURL, old.path != newURL.path {
            try? FileManager.default.removeItem(at: old)
        }
        return newURL
    }

    static func delete(theme: FontTheme) throws {
        guard let url = theme.fileURL else { return }
        try FileManager.default.removeItem(at: url)
    }

    private static func makeXML(name: String, headingFont: String, bodyFont: String) -> String {
        let eName = xmlEscape(name)
        let eHeading = xmlEscape(headingFont)
        let eBody = xmlEscape(bodyFont)
        return """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:fontScheme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="\(eName)">
          <a:majorFont>
            <a:latin typeface="\(eHeading)"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:majorFont>
          <a:minorFont>
            <a:latin typeface="\(eBody)"/>
            <a:ea typeface=""/>
            <a:cs typeface=""/>
          </a:minorFont>
        </a:fontScheme>
        """
    }

    private static func xmlEscape(_ s: String) -> String {
        s.replacingOccurrences(of: "&", with: "&amp;")
         .replacingOccurrences(of: "<", with: "&lt;")
         .replacingOccurrences(of: ">", with: "&gt;")
         .replacingOccurrences(of: "\"", with: "&quot;")
    }
}

// MARK: - Theme Row View

private class ThemeRowView: NSTableCellView {

    private let aaBox = NSBox()
    private let aaLabel = NSTextField(labelWithString: "Aa")
    private let nameLabel = NSTextField(labelWithString: "")
    private let headingLabel = NSTextField(labelWithString: "")
    private let bodyLabel = NSTextField(labelWithString: "")

    override init(frame: NSRect) {
        super.init(frame: frame)
        setup()
    }

    required init?(coder: NSCoder) {
        super.init(coder: coder)
        setup()
    }

    private func setup() {
        aaBox.boxType = .custom
        aaBox.borderWidth = 1
        aaBox.cornerRadius = 3
        aaBox.borderColor = .separatorColor
        aaBox.fillColor = .controlBackgroundColor
        aaBox.translatesAutoresizingMaskIntoConstraints = false

        aaLabel.alignment = .center
        aaLabel.translatesAutoresizingMaskIntoConstraints = false
        aaBox.addSubview(aaLabel)

        nameLabel.font = .systemFont(ofSize: 12, weight: .medium)
        nameLabel.translatesAutoresizingMaskIntoConstraints = false
        nameLabel.lineBreakMode = .byTruncatingTail

        headingLabel.font = .systemFont(ofSize: 10)
        headingLabel.textColor = .secondaryLabelColor
        headingLabel.translatesAutoresizingMaskIntoConstraints = false
        headingLabel.lineBreakMode = .byTruncatingTail

        bodyLabel.font = .systemFont(ofSize: 10)
        bodyLabel.textColor = .secondaryLabelColor
        bodyLabel.translatesAutoresizingMaskIntoConstraints = false
        bodyLabel.lineBreakMode = .byTruncatingTail

        let textStack = NSStackView(views: [nameLabel, headingLabel, bodyLabel])
        textStack.orientation = .vertical
        textStack.alignment = .leading
        textStack.spacing = 1
        textStack.translatesAutoresizingMaskIntoConstraints = false

        addSubview(aaBox)
        addSubview(textStack)

        NSLayoutConstraint.activate([
            aaBox.leadingAnchor.constraint(equalTo: leadingAnchor, constant: 10),
            aaBox.centerYAnchor.constraint(equalTo: centerYAnchor),
            aaBox.widthAnchor.constraint(equalToConstant: 42),
            aaBox.heightAnchor.constraint(equalToConstant: 42),

            aaLabel.centerXAnchor.constraint(equalTo: aaBox.centerXAnchor),
            aaLabel.centerYAnchor.constraint(equalTo: aaBox.centerYAnchor),

            textStack.leadingAnchor.constraint(equalTo: aaBox.trailingAnchor, constant: 10),
            textStack.trailingAnchor.constraint(equalTo: trailingAnchor, constant: -8),
            textStack.centerYAnchor.constraint(equalTo: centerYAnchor),
        ])
    }

    func configure(with theme: FontTheme) {
        let headingName = theme.headingFont.isEmpty ? NSFont.systemFont(ofSize: 18).fontName : theme.headingFont
        aaLabel.font = NSFont(name: headingName, size: 18) ?? .systemFont(ofSize: 18)
        nameLabel.stringValue = theme.name
        headingLabel.stringValue = theme.headingFont.isEmpty ? "(system)" : theme.headingFont
        bodyLabel.stringValue = theme.bodyFont.isEmpty ? "(system)" : theme.bodyFont
    }
}

// MARK: - Main View Controller

class ViewController: NSViewController {

    // Data
    private var themes: [FontTheme] = []
    private var selectedIndex: Int = -1
    private lazy var availableFontFamilies: [String] = NSFontManager.shared.availableFontFamilies.sorted()

    // Left panel
    private let tableView = NSTableView()
    private let scrollView = NSScrollView()
    private let addButton = NSButton()
    private let removeButton = NSButton()

    // Right panel
    private let nameField = NSTextField()
    private let headingPopUp = NSComboBox()
    private let bodyPopUp = NSComboBox()
    private let sampleBox = NSBox()
    private let headingSample = NSTextField(labelWithString: "Heading")
    private let bodySample = NSTextField(labelWithString: "Body text body text body text.\nBody text body text.")

    // MARK: - Lifecycle

    override func viewDidLoad() {
        super.viewDidLoad()
        buildUI()
        loadThemes()
    }

    override func viewDidAppear() {
        super.viewDidAppear()
        view.window?.title = "Couturier"
        view.window?.minSize = NSSize(width: 520, height: 360)
    }

    override var preferredContentSize: NSSize {
        get { NSSize(width: 620, height: 420) }
        set { }
    }

    // MARK: - UI Setup

    private func buildUI() {
        // Root split view
        let split = NSSplitView()
        split.isVertical = true
        split.dividerStyle = .thin
        split.translatesAutoresizingMaskIntoConstraints = false
        view.addSubview(split)
        NSLayoutConstraint.activate([
            split.topAnchor.constraint(equalTo: view.topAnchor),
            split.bottomAnchor.constraint(equalTo: view.bottomAnchor),
            split.leadingAnchor.constraint(equalTo: view.leadingAnchor),
            split.trailingAnchor.constraint(equalTo: view.trailingAnchor),
        ])

        split.addArrangedSubview(makeLeftPanel())
        split.addArrangedSubview(makeRightPanel())
        split.setHoldingPriority(.defaultLow, forSubviewAt: 0)
        split.setHoldingPriority(.defaultHigh, forSubviewAt: 1)
    }

    private func makeLeftPanel() -> NSView {
        let panel = NSView()
        panel.translatesAutoresizingMaskIntoConstraints = false
        panel.widthAnchor.constraint(equalToConstant: 200).isActive = true

        // Table
        let col = NSTableColumn(identifier: NSUserInterfaceItemIdentifier("col"))
        col.resizingMask = .autoresizingMask
        tableView.addTableColumn(col)
        tableView.headerView = nil
        tableView.rowHeight = 62
        tableView.intercellSpacing = NSSize(width: 0, height: 0)
        tableView.selectionHighlightStyle = .regular
        tableView.focusRingType = .none
        tableView.delegate = self
        tableView.dataSource = self

        scrollView.documentView = tableView
        scrollView.hasVerticalScroller = true
        scrollView.autohidesScrollers = true
        scrollView.borderType = .noBorder
        scrollView.translatesAutoresizingMaskIntoConstraints = false
        panel.addSubview(scrollView)

        // Bottom toolbar
        let toolbar = makeBottomToolbar()
        panel.addSubview(toolbar)

        NSLayoutConstraint.activate([
            scrollView.topAnchor.constraint(equalTo: panel.topAnchor),
            scrollView.leadingAnchor.constraint(equalTo: panel.leadingAnchor),
            scrollView.trailingAnchor.constraint(equalTo: panel.trailingAnchor),
            scrollView.bottomAnchor.constraint(equalTo: toolbar.topAnchor),

            toolbar.leadingAnchor.constraint(equalTo: panel.leadingAnchor),
            toolbar.trailingAnchor.constraint(equalTo: panel.trailingAnchor),
            toolbar.bottomAnchor.constraint(equalTo: panel.bottomAnchor),
            toolbar.heightAnchor.constraint(equalToConstant: 24),
        ])

        return panel
    }

    private func makeBottomToolbar() -> NSView {
        let bar = NSView()
        bar.translatesAutoresizingMaskIntoConstraints = false

        let sep = NSBox()
        sep.boxType = .separator
        sep.translatesAutoresizingMaskIntoConstraints = false
        bar.addSubview(sep)

        func toolbarButton(symbolName: String, fallbackName: String) -> NSButton {
            let btn = NSButton()
            btn.bezelStyle = .recessed
            btn.isBordered = false
            btn.imageScaling = .scaleProportionallyDown
            if let img = NSImage(systemSymbolName: symbolName, accessibilityDescription: nil) {
                btn.image = img
            } else {
                btn.image = NSImage(named: fallbackName)
            }
            btn.translatesAutoresizingMaskIntoConstraints = false
            return btn
        }

        removeButton.bezelStyle = .recessed
        removeButton.isBordered = false
        removeButton.imageScaling = .scaleProportionallyDown
        removeButton.image = NSImage(systemSymbolName: "minus", accessibilityDescription: "Remove")
            ?? NSImage(named: NSImage.removeTemplateName)
        removeButton.translatesAutoresizingMaskIntoConstraints = false
        removeButton.target = self
        removeButton.action = #selector(removeThemeTapped)
        removeButton.isEnabled = false

        addButton.bezelStyle = .recessed
        addButton.isBordered = false
        addButton.imageScaling = .scaleProportionallyDown
        addButton.image = NSImage(systemSymbolName: "plus", accessibilityDescription: "Add")
            ?? NSImage(named: NSImage.addTemplateName)
        addButton.translatesAutoresizingMaskIntoConstraints = false
        addButton.target = self
        addButton.action = #selector(addThemeTapped)

        bar.addSubview(removeButton)
        bar.addSubview(addButton)

        NSLayoutConstraint.activate([
            sep.topAnchor.constraint(equalTo: bar.topAnchor),
            sep.leadingAnchor.constraint(equalTo: bar.leadingAnchor),
            sep.trailingAnchor.constraint(equalTo: bar.trailingAnchor),

            removeButton.leadingAnchor.constraint(equalTo: bar.leadingAnchor, constant: 6),
            removeButton.centerYAnchor.constraint(equalTo: bar.centerYAnchor, constant: 1),
            removeButton.widthAnchor.constraint(equalToConstant: 20),
            removeButton.heightAnchor.constraint(equalToConstant: 20),

            addButton.leadingAnchor.constraint(equalTo: removeButton.trailingAnchor, constant: 2),
            addButton.centerYAnchor.constraint(equalTo: bar.centerYAnchor, constant: 1),
            addButton.widthAnchor.constraint(equalToConstant: 20),
            addButton.heightAnchor.constraint(equalToConstant: 20),
        ])

        return bar
    }

    private func makeRightPanel() -> NSView {
        let panel = NSView()
        panel.translatesAutoresizingMaskIntoConstraints = false

        // Name row
        let nameRowLabel = NSTextField(labelWithString: "Name:")
        nameRowLabel.font = .systemFont(ofSize: 13, weight: .medium)
        nameRowLabel.translatesAutoresizingMaskIntoConstraints = false

        nameField.placeholderString = "Theme name"
        nameField.font = .systemFont(ofSize: 13)
        nameField.translatesAutoresizingMaskIntoConstraints = false
        nameField.target = self
        nameField.action = #selector(nameFieldCommitted)
        (nameField.cell as? NSTextFieldCell)?.sendsActionOnEndEditing = true

        // Heading font
        let headingRowLabel = NSTextField(labelWithString: "Heading font:")
        headingRowLabel.font = .systemFont(ofSize: 12)
        headingRowLabel.textColor = .secondaryLabelColor
        headingRowLabel.translatesAutoresizingMaskIntoConstraints = false

        configureFontComboBox(headingPopUp)
        headingPopUp.target = self
        headingPopUp.action = #selector(headingFontChanged)
        headingPopUp.delegate = self

        // Body font
        let bodyRowLabel = NSTextField(labelWithString: "Body font:")
        bodyRowLabel.font = .systemFont(ofSize: 12)
        bodyRowLabel.textColor = .secondaryLabelColor
        bodyRowLabel.translatesAutoresizingMaskIntoConstraints = false

        configureFontComboBox(bodyPopUp)
        bodyPopUp.target = self
        bodyPopUp.action = #selector(bodyFontChanged)
        bodyPopUp.delegate = self

        // Sample box
        sampleBox.title = "Sample"
        sampleBox.titleFont = .systemFont(ofSize: 11)
        sampleBox.translatesAutoresizingMaskIntoConstraints = false

        headingSample.font = .systemFont(ofSize: 24)
        headingSample.translatesAutoresizingMaskIntoConstraints = false

        bodySample.font = .systemFont(ofSize: 12)
        bodySample.maximumNumberOfLines = 0
        bodySample.translatesAutoresizingMaskIntoConstraints = false

        let sampleStack = NSStackView(views: [headingSample, bodySample])
        sampleStack.orientation = .vertical
        sampleStack.alignment = .leading
        sampleStack.spacing = 6
        sampleStack.translatesAutoresizingMaskIntoConstraints = false
        sampleBox.addSubview(sampleStack)

        NSLayoutConstraint.activate([
            sampleStack.topAnchor.constraint(equalTo: sampleBox.topAnchor, constant: 18),
            sampleStack.leadingAnchor.constraint(equalTo: sampleBox.leadingAnchor, constant: 12),
            sampleStack.trailingAnchor.constraint(equalTo: sampleBox.trailingAnchor, constant: -12),
            sampleStack.bottomAnchor.constraint(equalTo: sampleBox.bottomAnchor, constant: -12),
        ])

        // Layout all into panel
        [nameRowLabel, nameField,
         headingRowLabel, headingPopUp,
         bodyRowLabel, bodyPopUp,
         sampleBox].forEach { panel.addSubview($0) }

        NSLayoutConstraint.activate([
            // Name row
            nameRowLabel.topAnchor.constraint(equalTo: panel.topAnchor, constant: 20),
            nameRowLabel.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            nameRowLabel.widthAnchor.constraint(equalToConstant: 50),
            nameRowLabel.centerYAnchor.constraint(equalTo: nameField.centerYAnchor),

            nameField.topAnchor.constraint(equalTo: panel.topAnchor, constant: 20),
            nameField.leadingAnchor.constraint(equalTo: nameRowLabel.trailingAnchor, constant: 8),
            nameField.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),

            // Heading
            headingRowLabel.topAnchor.constraint(equalTo: nameField.bottomAnchor, constant: 16),
            headingRowLabel.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            headingRowLabel.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),

            headingPopUp.topAnchor.constraint(equalTo: headingRowLabel.bottomAnchor, constant: 4),
            headingPopUp.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            headingPopUp.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),

            // Body
            bodyRowLabel.topAnchor.constraint(equalTo: headingPopUp.bottomAnchor, constant: 12),
            bodyRowLabel.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            bodyRowLabel.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),

            bodyPopUp.topAnchor.constraint(equalTo: bodyRowLabel.bottomAnchor, constant: 4),
            bodyPopUp.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            bodyPopUp.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),

            // Sample
            sampleBox.topAnchor.constraint(equalTo: bodyPopUp.bottomAnchor, constant: 16),
            sampleBox.leadingAnchor.constraint(equalTo: panel.leadingAnchor, constant: 20),
            sampleBox.trailingAnchor.constraint(equalTo: panel.trailingAnchor, constant: -20),
            sampleBox.bottomAnchor.constraint(lessThanOrEqualTo: panel.bottomAnchor, constant: -20),
        ])

        setDetailEnabled(false)
        return panel
    }

    private func configureFontComboBox(_ cb: NSComboBox) {
        cb.translatesAutoresizingMaskIntoConstraints = false
        cb.completes = true
        cb.hasVerticalScroller = true
        cb.numberOfVisibleItems = 10
        cb.addItems(withObjectValues: availableFontFamilies)
    }

    // MARK: - Data

    private func loadThemes() {
        themes = FontThemeManager.loadThemes()
        tableView.reloadData()
        if !themes.isEmpty {
            tableView.selectRowIndexes(IndexSet(integer: 0), byExtendingSelection: false)
        }
    }

    private func showTheme(_ theme: FontTheme) {
        nameField.stringValue = theme.name
        selectFont(headingPopUp, name: theme.headingFont)
        selectFont(bodyPopUp, name: theme.bodyFont)
        refreshSample()
        setDetailEnabled(true)
    }

    private func selectFont(_ cb: NSComboBox, name: String) {
        if name.isEmpty {
            cb.stringValue = availableFontFamilies.first ?? ""
        } else if availableFontFamilies.contains(name) {
            cb.selectItem(withObjectValue: name)
        } else {
            // Font not installed — show the name but don't crash
            cb.stringValue = name
        }
    }

    private func setDetailEnabled(_ enabled: Bool) {
        nameField.isEnabled = enabled
        headingPopUp.isEnabled = enabled
        bodyPopUp.isEnabled = enabled
        if !enabled {
            nameField.stringValue = ""
            headingPopUp.stringValue = ""
            bodyPopUp.stringValue = ""
            resetSampleFonts()
        }
    }

    private func refreshSample() {
        let h = headingPopUp.stringValue
        let b = bodyPopUp.stringValue
        headingSample.font = NSFont(name: h, size: 24) ?? .systemFont(ofSize: 24)
        bodySample.font = NSFont(name: b, size: 12) ?? .systemFont(ofSize: 12)
    }

    private func resetSampleFonts() {
        headingSample.font = .systemFont(ofSize: 24)
        bodySample.font = .systemFont(ofSize: 12)
    }

    private func saveCurrentTheme() {
        guard selectedIndex >= 0, selectedIndex < themes.count else { return }
        do {
            let url = try FontThemeManager.save(theme: themes[selectedIndex])
            themes[selectedIndex].fileURL = url
        } catch {
            presentError(error)
        }
    }

    // MARK: - Actions

    @objc private func nameFieldCommitted() {
        guard selectedIndex >= 0, selectedIndex < themes.count else { return }
        let newName = nameField.stringValue.trimmingCharacters(in: .whitespacesAndNewlines)
        guard !newName.isEmpty, newName != themes[selectedIndex].name else { return }

        // Ensure unique name
        guard !themes.enumerated().contains(where: { $0.offset != selectedIndex && $0.element.name == newName }) else {
            let alert = NSAlert()
            alert.messageText = "Name already in use"
            alert.informativeText = "A theme named \"\(newName)\" already exists."
            alert.runModal()
            nameField.stringValue = themes[selectedIndex].name
            return
        }

        do {
            let newURL = try FontThemeManager.rename(theme: &themes[selectedIndex], newName: newName)
            themes[selectedIndex].fileURL = newURL
            // Re-sort
            let theme = themes[selectedIndex]
            themes.sort { $0.name.localizedCaseInsensitiveCompare($1.name) == .orderedAscending }
            tableView.reloadData()
            if let newIdx = themes.firstIndex(where: { $0.fileURL == theme.fileURL }) {
                selectedIndex = newIdx
                tableView.selectRowIndexes(IndexSet(integer: newIdx), byExtendingSelection: false)
            }
        } catch {
            presentError(error)
        }
    }

    @objc private func headingFontChanged() {
        guard selectedIndex >= 0, selectedIndex < themes.count else { return }
        themes[selectedIndex].headingFont = headingPopUp.stringValue
        saveCurrentTheme()
        refreshSample()
        tableView.reloadData(forRowIndexes: IndexSet(integer: selectedIndex), columnIndexes: IndexSet(integer: 0))
    }

    @objc private func bodyFontChanged() {
        guard selectedIndex >= 0, selectedIndex < themes.count else { return }
        themes[selectedIndex].bodyFont = bodyPopUp.stringValue
        saveCurrentTheme()
        refreshSample()
        tableView.reloadData(forRowIndexes: IndexSet(integer: selectedIndex), columnIndexes: IndexSet(integer: 0))
    }

    @objc private func addThemeTapped() {
        // Pick a unique default name
        var baseName = "Custom"
        var counter = 2
        while themes.contains(where: { $0.name == baseName }) {
            baseName = "Custom \(counter)"
            counter += 1
        }
        let defaultFont = availableFontFamilies.first ?? "Helvetica Neue"
        var theme = FontTheme(name: baseName, headingFont: defaultFont, bodyFont: defaultFont)
        do {
            let url = try FontThemeManager.save(theme: theme)
            theme.fileURL = url
            themes.append(theme)
            themes.sort { $0.name.localizedCaseInsensitiveCompare($1.name) == .orderedAscending }
            tableView.reloadData()
            if let idx = themes.firstIndex(where: { $0.fileURL == url }) {
                tableView.selectRowIndexes(IndexSet(integer: idx), byExtendingSelection: false)
                tableView.scrollRowToVisible(idx)
                // Focus name field for immediate renaming
                view.window?.makeFirstResponder(nameField)
                nameField.selectText(nil)
            }
        } catch {
            presentError(error)
        }
    }

    @objc private func removeThemeTapped() {
        guard selectedIndex >= 0, selectedIndex < themes.count else { return }
        let theme = themes[selectedIndex]
        let alert = NSAlert()
        alert.messageText = "Delete \"\(theme.name)\"?"
        alert.informativeText = "This will permanently remove the font theme file."
        alert.addButton(withTitle: "Delete")
        alert.addButton(withTitle: "Cancel")
        alert.alertStyle = .warning
        guard let window = view.window else { return }
        alert.beginSheetModal(for: window) { [weak self] response in
            guard let self, response == .alertFirstButtonReturn else { return }
            do {
                try FontThemeManager.delete(theme: theme)
                self.themes.remove(at: self.selectedIndex)
                self.tableView.reloadData()
                if self.themes.isEmpty {
                    self.selectedIndex = -1
                    self.setDetailEnabled(false)
                    self.removeButton.isEnabled = false
                } else {
                    let idx = min(self.selectedIndex, self.themes.count - 1)
                    self.tableView.selectRowIndexes(IndexSet(integer: idx), byExtendingSelection: false)
                }
            } catch {
                self.presentError(error)
            }
        }
    }
}

// MARK: - NSTableViewDataSource

extension ViewController: NSTableViewDataSource {
    func numberOfRows(in tableView: NSTableView) -> Int { themes.count }
}

// MARK: - NSTableViewDelegate

extension ViewController: NSTableViewDelegate {
    func tableView(_ tableView: NSTableView, viewFor tableColumn: NSTableColumn?, row: Int) -> NSView? {
        let id = NSUserInterfaceItemIdentifier("ThemeRow")
        let cell = (tableView.makeView(withIdentifier: id, owner: self) as? ThemeRowView) ?? {
            let v = ThemeRowView()
            v.identifier = id
            return v
        }()
        cell.configure(with: themes[row])
        return cell
    }

    func tableViewSelectionDidChange(_ notification: Notification) {
        selectedIndex = tableView.selectedRow
        removeButton.isEnabled = selectedIndex >= 0
        if selectedIndex >= 0, selectedIndex < themes.count {
            showTheme(themes[selectedIndex])
        } else {
            setDetailEnabled(false)
        }
    }
}

// MARK: - NSComboBoxDelegate

extension ViewController: NSComboBoxDelegate {
    func comboBoxSelectionDidChange(_ notification: Notification) {
        guard let cb = notification.object as? NSComboBox else { return }
        if cb === headingPopUp { headingFontChanged() }
        else if cb === bodyPopUp { bodyFontChanged() }
    }
}
