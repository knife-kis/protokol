package ru.citlab24.protokol.tabs.buildingTab;

import javafx.application.Platform;
import javafx.embed.swing.JFXPanel;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.dialogs.*;
import ru.citlab24.protokol.tabs.modules.lighting.LightingTab;
import ru.citlab24.protokol.tabs.modules.lighting.ArtificialLightingTab;
import ru.citlab24.protokol.tabs.modules.med.RadiationTab;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationTab;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.RoomListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;


import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.event.*;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BuildingTab extends JPanel {
    private static final Logger logger = LoggerFactory.getLogger(BuildingTab.class);
    // Константы для цветов и стилей
    private static final java.awt.Color FLOOR_PANEL_COLOR = new java.awt.Color(0, 115, 200);
    private static final java.awt.Color SPACE_PANEL_COLOR = new java.awt.Color(76, 175, 80);
    private static final java.awt.Color ROOM_PANEL_COLOR  = new java.awt.Color(156, 39, 176);

    private static final Font HEADER_FONT =
            UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);
    private static final Dimension BUTTON_PANEL_SIZE = new Dimension(5, 5);

    private Building building;
    private BuildingModelOps ops;
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private JList<Section> sectionList;

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;
    private JTextField projectNameField;
    // === Фильтр типов помещений (Жилые / Офисные / Общественные) ===
    private JToggleButton filterApartmentBtn;
    private JToggleButton filterOfficeBtn;
    private JToggleButton filterPublicBtn;

    public BuildingTab(Building building) {
        // 1) сохраняем в поле, 2) создаём дефолт, если пришёл null
        this.building = (building != null) ? building : new Building();
        if (this.building.getSections().isEmpty()) {
            this.building.addSection(new Section("Секция 1", 0));
        }
        this.ops = new BuildingModelOps(this.building); // ← добавили
        initComponents();
    }

    private static String normalizeRoomBaseName(String name) {
        return BuildingModelOps.normalizeRoomBaseName(name);
    }

    private void initComponents() {
        setLayout(new BorderLayout());
        add(createProjectNamePanel(), BorderLayout.NORTH);
        add(createBuildingPanel(), BorderLayout.CENTER);
        add(createActionButtons(), BorderLayout.SOUTH);

        setupListSelectionListeners();
    }

    private void setupListSelectionListeners() {
        // Инициализация будет после создания компонентов
    }

    private JPanel createProjectNamePanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(0, 10, 0, 10));

        JLabel label = new JLabel("Название проекта:");
        label.setFont(HEADER_FONT);

        projectNameField = new JTextField(30);
        projectNameField.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        String bName = (this.building.getName() != null) ? this.building.getName() : "";
        projectNameField.setText(bName);

        panel.add(label);
        panel.add(projectNameField);

        return panel;
    }

    private JPanel createBuildingPanel() {
        JPanel panel = new JPanel(new GridLayout(1, 3, 15, 15));
        panel.add(createSectionsAndFloorsPanel());
        panel.add(createSpacesPanelWithFilter()); // ← ТУТ теперь панель с фильтром
        panel.add(createListPanel("Комнаты в помещении", ROOM_PANEL_COLOR, roomListModel,
                new RoomListRenderer(), this::createRoomButtons));
        return panel;
    }
    /** Средняя колонка: «Помещения на этаже» + панель фильтров (Жилые/Офисные/Общественные). */
    private JPanel createSpacesPanelWithFilter() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения на этаже", SPACE_PANEL_COLOR));

        // ===== Список помещений =====
        spaceList = new JList<>(spaceListModel);
        spaceList.setCellRenderer(new SpaceListRenderer());
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.setDragEnabled(true);
        spaceList.setDropMode(DropMode.INSERT);
        spaceList.setTransferHandler(spaceReorderHandler);

        // ===== Фильтры =====
        JPanel filters = buildSpaceFilterPanel();

        panel.add(filters, BorderLayout.NORTH);
        panel.add(new JScrollPane(spaceList), BorderLayout.CENTER);
        panel.add(createSpaceButtons(), BorderLayout.SOUTH);
        return panel;
    }
    //** Верхняя панель с 3 переключателями-фильтрами. Любая комбинация допустима. */
    private JPanel buildSpaceFilterPanel() {
        JPanel wrap = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 6));

        filterApartmentBtn = new JToggleButton("Жилые");
        filterOfficeBtn    = new JToggleButton("Офисные");
        filterPublicBtn    = new JToggleButton("Общественные");

        // По умолчанию показываем всё
        filterApartmentBtn.setSelected(true);
        filterOfficeBtn.setSelected(true);
        filterPublicBtn.setSelected(true);

        // Стиль
        java.util.List<JToggleButton> all = java.util.List.of(filterApartmentBtn, filterOfficeBtn, filterPublicBtn);
        for (JToggleButton b : all) {
            b.setFocusPainted(false);
            b.setFont(new Font("Segoe UI", Font.PLAIN, 13));
        }

        // Любое изменение тумблеров → пересобрать список этажей и помещений
        java.awt.event.ItemListener l = e -> {
            refreshFloorListForSelectedSection(); // отфильтруем этажи
            updateSpaceList();                    // и сами помещения на выбранном этаже
        };
        filterApartmentBtn.addItemListener(l);
        filterOfficeBtn.addItemListener(l);
        filterPublicBtn.addItemListener(l);

        wrap.add(new JLabel("Фильтр:"));
        wrap.add(filterApartmentBtn);
        wrap.add(filterOfficeBtn);
        wrap.add(filterPublicBtn);
        return wrap;
    }

    private <T> JPanel createListPanel(String title, Color color, ListModel<T> model,
                                       ListCellRenderer<? super T> renderer,
                                       Supplier<JPanel> buttonPanelSupplier) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder(title, color));

        JList<T> list = new JList<>(model);
        list.setCellRenderer(renderer);
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Сохраняем ссылки на списки + включаем DnD там, где нужно
        if ("Этажи здания".equals(title)) {
            floorList = (JList<Floor>) list;
        } else if ("Помещения на этаже".equals(title)) {
            spaceList = (JList<Space>) list;
            spaceList.setDragEnabled(true);
            spaceList.setDropMode(DropMode.INSERT);
            spaceList.setTransferHandler(spaceReorderHandler);
        } else if ("Комнаты в помещении".equals(title)) {
            roomList = (JList<Room>) list;
            roomList.setDragEnabled(true);
            roomList.setDropMode(DropMode.INSERT);
            roomList.setTransferHandler(roomReorderHandler);
        }

        panel.add(new JScrollPane(list), BorderLayout.CENTER);
        panel.add(buttonPanelSupplier.get(), BorderLayout.SOUTH);

        return panel;
    }


    private TitledBorder createTitledBorder(String title, Color color) {
        return BorderFactory.createTitledBorder(
                null, title, TitledBorder.LEFT, TitledBorder.TOP, HEADER_FONT, color);
    }

    private JPanel createFloorButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.LAYER_GROUP, new Color(63,81,181), this::manageSections),
                createStyledButton("", FontAwesomeSolid.COPY,        new Color(0,150,136), this::copySection), // ← НОВОЕ
                createStyledButton("", FontAwesomeSolid.PLUS,        new Color(46,125,50),  this::addFloor),
                createStyledButton("", FontAwesomeSolid.CLONE,       new Color(100,181,246), this::copyFloor),
                createStyledButton("", FontAwesomeSolid.EDIT,        new Color(255,152,0),  this::editFloor),
                createStyledButton("", FontAwesomeSolid.TRASH,       new Color(198,40,40),  this::removeFloor)
        );

    }

    private JPanel createSpaceButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.PLUS, new Color(46, 125, 50), this::addSpace),
                createStyledButton("", FontAwesomeSolid.CLONE, new Color(100, 181, 246), this::copySpace),
                createStyledButton("", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editSpace),
                createStyledButton("", FontAwesomeSolid.TRASH, new Color(198, 40, 40), this::removeSpace)
        );
    }

    private JPanel createRoomButtons() {
        return createButtonPanel(
                createStyledButton("", FontAwesomeSolid.PLUS, new Color(46, 125, 50), this::addRoom),
                createStyledButton("", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editRoom),
                createStyledButton("", FontAwesomeSolid.TRASH, new Color(198, 40, 40), this::removeRoom)
        );
    }

    private JPanel createButtonPanel(JButton... buttons) {
        JPanel panel = new JPanel(new GridLayout(1, buttons.length, BUTTON_PANEL_SIZE.width, BUTTON_PANEL_SIZE.height));
        for (JButton button : buttons) {
            panel.add(button);
        }
        return panel;
    }

    private JPanel createActionButtons() {
        JPanel wrap = new JPanel(new BorderLayout());
        final JFXPanel fx = new JFXPanel();
        wrap.add(fx, BorderLayout.CENTER);

        Platform.runLater(() -> {
            // Кнопки
            javafx.scene.control.Button btnLoad    = new javafx.scene.control.Button("Загрузить проект");
            javafx.scene.control.Button btnSave    = new javafx.scene.control.Button("Сохранить проект");
            javafx.scene.control.Button btnCalc    = new javafx.scene.control.Button("Рассчитать показатели");
            javafx.scene.control.Button btnSummary = new javafx.scene.control.Button("Сводка квартир");
            javafx.scene.control.Button btnExport  = new javafx.scene.control.Button("Экспорт: все модули (одной книгой)");

            // Иконки (Ikonli JavaFX)
            org.kordamp.ikonli.javafx.FontIcon icLoad    = new org.kordamp.ikonli.javafx.FontIcon("fas-folder-open");
            org.kordamp.ikonli.javafx.FontIcon icSave    = new org.kordamp.ikonli.javafx.FontIcon("fas-save");
            org.kordamp.ikonli.javafx.FontIcon icCalc    = new org.kordamp.ikonli.javafx.FontIcon("fas-calculator");
            org.kordamp.ikonli.javafx.FontIcon icSummary = new org.kordamp.ikonli.javafx.FontIcon("fas-table");
            org.kordamp.ikonli.javafx.FontIcon icExport  = new org.kordamp.ikonli.javafx.FontIcon("fas-file-excel");
            icLoad.setIconSize(16); icSave.setIconSize(16); icCalc.setIconSize(16); icSummary.setIconSize(16); icExport.setIconSize(16);
            btnLoad.setGraphic(icLoad); btnSave.setGraphic(icSave); btnCalc.setGraphic(icCalc); btnSummary.setGraphic(icSummary); btnExport.setGraphic(icExport);

            // CSS-классы для цветов/hover
            btnLoad.getStyleClass().addAll("button", "btn-load");
            btnSave.getStyleClass().addAll("button", "btn-save");
            btnCalc.getStyleClass().addAll("button", "btn-calc");
            btnSummary.getStyleClass().addAll("button", "btn-summary");
            btnExport.getStyleClass().addAll("button", "btn-export");

            // Растягиваем равномерно
            javafx.scene.layout.HBox box = new javafx.scene.layout.HBox(10, btnLoad, btnSave, btnCalc, btnSummary, btnExport);
            box.getStyleClass().addAll("controls-bar", "theme-light");
            for (javafx.scene.control.Button b : java.util.List.of(btnLoad, btnSave, btnCalc, btnSummary, btnExport)) {
                b.setMaxWidth(Double.MAX_VALUE);
                javafx.scene.layout.HBox.setHgrow(b, javafx.scene.layout.Priority.ALWAYS);
            }

            // Сцена + CSS
            javafx.scene.Scene scene = new javafx.scene.Scene(box);
            scene.getStylesheets().add(
                    java.util.Objects.requireNonNull(getClass().getResource("/ui/protokol.css")).toExternalForm()
            );
            fx.setScene(scene);

            // Горячая смена темы (Ctrl+T)
            scene.setOnKeyPressed(ke -> {
                if (ke.isControlDown() && ke.getCode() == javafx.scene.input.KeyCode.T) {
                    java.util.List<String> cls = box.getStyleClass();
                    if (cls.contains("theme-light")) { cls.remove("theme-light"); cls.add("theme-dark"); }
                    else { cls.remove("theme-dark"); cls.add("theme-light"); }
                }
            });

            // Цвета и "активность"
            final String COL_LOAD    = "#3949ab";
            final String COL_SAVE    = "#43a047";
            final String COL_CALC    = "#1e88e5";
            final String COL_SUMMARY = "#00897b";
            final String COL_EXPORT  = "#ef6c00";

            final java.util.List<javafx.scene.control.Button> all =
                    java.util.List.of(btnLoad, btnSave, btnCalc, btnSummary, btnExport);

            final java.util.function.BiConsumer<javafx.scene.control.Button, String> markActive = (btn, hex) -> {
                all.forEach(ru.citlab24.protokol.ui.fx.ActiveGlow::clear);
                // неон на ~50%
                ru.citlab24.protokol.ui.fx.ActiveGlow.setActive(btn, hex, 0.5);
            };

            // Обработчики: сначала отметить активной, потом вызвать Swing-логику
            btnLoad.setOnAction(ev -> {
                markActive.accept(btnLoad, COL_LOAD);
                javax.swing.SwingUtilities.invokeLater(() -> loadProject(null));
            });
            btnSave.setOnAction(ev -> {
                markActive.accept(btnSave, COL_SAVE);
                javax.swing.SwingUtilities.invokeLater(() -> saveProject(null));
            });
            btnCalc.setOnAction(ev -> {
                markActive.accept(btnCalc, COL_CALC);
                javax.swing.SwingUtilities.invokeLater(() -> calculateMetrics(null));
            });
            btnSummary.setOnAction(ev -> {
                markActive.accept(btnSummary, COL_SUMMARY);
                javax.swing.SwingUtilities.invokeLater(() -> showApartmentSummary(null));
            });
            btnExport.setOnAction(ev -> {
                markActive.accept(btnExport, COL_EXPORT);
                javax.swing.SwingUtilities.invokeLater(() -> {
                    // фиксация редактирования и экспорт (твой код как был)
                    KeyboardFocusManager kfm = KeyboardFocusManager.getCurrentKeyboardFocusManager();
                    Component fo = (kfm != null) ? kfm.getFocusOwner() : null;
                    JTable editingTable = (fo == null) ? null
                            : (JTable) SwingUtilities.getAncestorOfClass(JTable.class, fo);
                    if (editingTable != null && editingTable.isEditing()) {
                        try { editingTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
                    }
                    RadiationTab rt = getRadiationTab();
                    if (rt != null) rt.updateRoomSelectionStates();
                    LightingTab lt = getLightingTab();
                    if (lt != null) lt.updateRoomSelectionStates();
                    MicroclimateTab mt = getMicroclimateTab();
                    if (mt != null) mt.updateRoomSelectionStates();

                    Window w = SwingUtilities.getWindowAncestor(this);
                    MainFrame frame = (w instanceof MainFrame) ? (MainFrame) w : null;
                    ru.citlab24.protokol.export.AllExcelExporter.exportAll(frame, building, this);
                });
            });
        });

        return wrap;
    }

    private void copySection(ActionEvent e) {
        if (sectionList == null || sectionList.isSelectionEmpty()) {
            showMessage("Выберите секцию для копирования", "Информация", JOptionPane.INFORMATION_MESSAGE);
            return;
        }
        // 0) Сначала зафиксируем состояния из вкладок в модель
        MicroclimateTab microTab = getMicroclimateTab();
        if (microTab != null) microTab.updateRoomSelectionStates(); // ← МИКРОКЛИМАТ в Room

        // 1) Сохраняем «снимки» галочек
        Map<String, Boolean> savedMicroSelections = saveMicroclimateSelections(); // МИКРОКЛИМАТ
        Map<String, Boolean> savedRadSelections   = saveRadiationSelections();    // РАДИАЦИЯ (как было)

        int srcIdx = sectionList.getSelectedIndex();
        Section src = sectionListModel.get(srcIdx);

        String suggested = generateUniqueSectionName(src.getName());
        String newName = showInputDialog("Копировать секцию", "Название новой секции:", suggested);
        if (newName == null || newName.trim().isEmpty()) return;

        // создаём новую секцию в конце списка
        int newPosition = building.getSections().size();
        Section dst = new Section();
        dst.setName(newName.trim());
        dst.setPosition(newPosition);
        building.addSection(dst);

        int dstIdx = newPosition;

        // 2) Глубоко копируем этажи исходной секции (rooms с НОВЫМИ id)
        java.util.List<Floor> toAdd = new java.util.ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == srcIdx) {
                Floor fCopy = createFloorCopy(f);
                fCopy.setSectionIndex(dstIdx);
                toAdd.add(fCopy);
            }
        }
        for (Floor fCopy : toAdd) building.addFloor(fCopy);

        // 3) ВОССТАНОВЛЕНИЕ МИКРОКЛИМАТА: применяем снимок к ВСЕМ комнатам здания
        restoreMicroclimateSelections(savedMicroSelections);

        // 4) UI
        refreshSectionListModel();
        sectionList.setSelectedIndex(dstIdx);
        refreshFloorListForSelectedSection();
        updateSpaceList();
        updateRoomList();

        // 5) Обновляем вкладки (порядок важен)
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        restoreRadiationSelections(savedRadSelections);

        updateLightingTab(building, /*autoApplyDefaults=*/true);

        if (microTab != null) {
            try { microTab.selectSectionByIndex(dstIdx); } catch (Throwable ignore) {}
        }

        showMessage("Секция «" + src.getName() + "» успешно скопирована в «" + newName + "».",
                "Готово", JOptionPane.INFORMATION_MESSAGE);
    }

    private String generateUniqueSectionName(String base) {
        if (base == null || base.isEmpty()) base = "Секция";
        String plain = base + " (копия)";
        java.util.Set<String> names = new java.util.HashSet<>();
        for (Section s : building.getSections()) {
            if (s.getName() != null) names.add(s.getName());
        }
        if (!names.contains(plain)) return plain;

        // если имя занято — добавляем счётчик
        int n = 2;
        while (names.contains(plain + " " + n)) n++;
        return plain + " " + n;
    }

    private JPanel createSectionsAndFloorsPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Секции и этажи", FLOOR_PANEL_COLOR));

        // СЛЕВА — список секций
        sectionList = new JList<>(sectionListModel);
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setFixedCellHeight(28);
        refreshSectionListModel();

// СПРАВА — список этажей текущей секции
        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new FloorListRenderer());
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setFixedCellHeight(28);

        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
        refreshFloorListForSelectedSection();

        sectionList.setDragEnabled(true);
        sectionList.setDropMode(DropMode.INSERT);
        sectionList.setTransferHandler(sectionReorderHandler);

        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);
        floorList.setDragEnabled(true);
        floorList.setDropMode(DropMode.INSERT);
        floorList.setTransferHandler(floorReorderHandler);

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloorListForSelectedSection();
        });

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(sectionList), new JScrollPane(floorList));
        split.setResizeWeight(0.35);
        panel.add(split, BorderLayout.CENTER);
        panel.add(createFloorButtons(), BorderLayout.SOUTH);
        return panel;
    }
    // Флажки, чтобы понимать откуда тянем (этажи → на секцию, или секции ↔ секции)
    private boolean draggingFromFloor = false;

    // Перестановка этажей ВНУТРИ текущей секции
    private final TransferHandler floorReorderHandler = new ReorderHandler<Floor>() {
        @Override protected Transferable createTransferable(JComponent c) {
            draggingFromFloor = true;                    // помечаем источник DnD
            return super.createTransferable(c);
        }
        @Override protected void afterReorder(DefaultListModel<Floor> model) {
            // Переименовываем позиции по текущему порядку в списке
            for (int i = 0; i < model.size(); i++) {
                model.get(i).setPosition(i);
            }
            draggingFromFloor = false;
        }
    };

    // Перестановка секций МЕЖДУ собой (ПЕРЕНОС ЭТАЖЕЙ НА СЕКЦИЮ ЗАПРЕЩЁН)
    private final TransferHandler sectionReorderHandler = new ReorderHandler<Section>() {

        @Override
        public boolean canImport(TransferSupport s) {
            // Запрещаем бросать ЭТАЖИ на список секций (перенос между секциями)
            if (draggingFromFloor) return false;
            return super.canImport(s);
        }

        @Override
        public boolean importData(TransferSupport s) {
            // Перенос этажа на секцию запрещён
            if (draggingFromFloor) return false;
            // Разрешаем только перестановку СЕКЦИЙ между собой
            return super.importData(s);
        }

        @Override
        protected void afterReorder(DefaultListModel<Section> model) {
            // 1) Текущий выбор секции (чтобы восстановить выделение)
            Section selected = (sectionList != null) ? sectionList.getSelectedValue() : null;

            // 2) Снимок старого списка секций (для ремапа индексов этажей)
            java.util.List<Section> oldSections = new java.util.ArrayList<>(building.getSections());

            // 3) Собираем новый порядок секций из модели JList и обновляем position
            java.util.List<Section> newSections = new java.util.ArrayList<>();
            for (int i = 0; i < model.size(); i++) {
                Section s = model.get(i);
                s.setPosition(i);
                newSections.add(s);
            }
            // 4) Карта: ссылка на секцию → новый индекс
            java.util.Map<Section, Integer> newIndexByRef = new java.util.HashMap<>();
            for (int i = 0; i < newSections.size(); i++) {
                newIndexByRef.put(newSections.get(i), i);
            }
            // 5) Ремапим sectionIndex у этажей так, чтобы они остались в СВОИХ секциях
            for (Floor f : building.getFloors()) {
                int oldIdx = f.getSectionIndex();
                if (oldIdx >= 0 && oldIdx < oldSections.size()) {
                    Section oldSec = oldSections.get(oldIdx);
                    Integer ni = newIndexByRef.get(oldSec);
                    f.setSectionIndex((ni != null) ? ni : 0);
                } else {
                    f.setSectionIndex(0);
                }
            }
            // 6) Фиксируем новые секции в модели здания и обновляем UI
            building.setSections(newSections);
            refreshSectionListModel();
            if (selected != null) {
                int selIdx = building.getSections().indexOf(selected);
                sectionList.setSelectedIndex(selIdx >= 0 ? selIdx : 0);
            } else if (!sectionListModel.isEmpty()) {
                sectionList.setSelectedIndex(0);
            }
            refreshFloorListForSelectedSection();
        }
    };

    private void refreshSectionListModel() {
        sectionListModel.clear();
        for (Section s : building.getSections()) sectionListModel.addElement(s);
    }

    private void refreshFloorListForSelectedSection() {
        floorListModel.clear();
        if (sectionList == null) return;

        int secIdx = Math.max(0, sectionList.getSelectedIndex());

        // Снимем текущий выбор, чтобы попытаться восстановить
        Floor previouslySelected = (floorList != null) ? floorList.getSelectedValue() : null;

        java.util.List<Floor> all = new java.util.ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == secIdx) all.add(f);
        }
        all.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        // Состояние фильтров
        boolean filtersReady = (filterApartmentBtn != null && filterOfficeBtn != null && filterPublicBtn != null);

        for (Floor f : all) {
            boolean include;
            if (!filtersReady) {
                include = true; // до инициализации фильтров ведём себя как раньше
            } else if (f.getSpaces().isEmpty()) {
                include = true; // пустые этажи всегда видны
            } else {
                include = false;
                for (Space s : f.getSpaces()) {
                    if (isSpaceVisibleByFilter(s)) { include = true; break; }
                }
            }
            if (include) floorListModel.addElement(f);
        }

        // Восстановление/установка выделения
        if (!floorListModel.isEmpty()) {
            int idx = (previouslySelected != null) ? floorListModel.indexOf(previouslySelected) : -1;
            floorList.setSelectedIndex(idx >= 0 ? idx : 0);
        } else {
            // Если фильтр скрыл все этажи секции — очистим зависимые списки
            if (spaceListModel != null) spaceListModel.clear();
            if (roomListModel  != null) roomListModel.clear();
        }
    }

    private JButton createStyledButton(String text, FontAwesomeSolid icon, java.awt.Color bgColor, ActionListener action) {
        JButton btn = new JButton(text, FontIcon.of(icon, 16, java.awt.Color.WHITE));
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setFont(UIManager.getFont("Button.font"));
        btn.addActionListener(action);

        btn.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                btn.setBackground(bgColor.darker());
            }

            @Override
            public void mouseExited(MouseEvent e) {
                btn.setBackground(bgColor);
            }
        });
        return btn;
    }

    void manageSections(ActionEvent e) {
        JFrame frame = (JFrame) SwingUtilities.getWindowAncestor(this);
        ManageSectionsDialog dlg = new ManageSectionsDialog(frame, building.getSections());
        if (!dlg.showDialog()) return;

        // 1) Запоминаем старые секции (для ремапа этажей по имени)
        List<Section> oldSections = new ArrayList<>(building.getSections());
        List<Section> updated = dlg.getSections();

        // 2) Строим карту "имя секции -> новый индекс"
        Map<String, Integer> nameToNewIndex = new HashMap<>();
        for (int i = 0; i < updated.size(); i++) {
            String name = updated.get(i).getName();
            if (name != null) nameToNewIndex.put(name, i);
        }

        // 3) Ремапим sectionIndex у всех этажей по ИМЕНИ старой секции.
        // Если имя не найдено в новом списке — отправляем этаж в секцию 0.
        for (Floor f : building.getFloors()) {
            int oldIdx = f.getSectionIndex();
            String oldName = (oldIdx >= 0 && oldIdx < oldSections.size())
                    ? oldSections.get(oldIdx).getName()
                    : null;
            Integer newIdx = (oldName == null) ? null : nameToNewIndex.get(oldName);
            f.setSectionIndex(newIdx != null ? newIdx : 0);
        }

        // 4) Фиксируем новые секции и обновляем UI
        building.setSections(updated);
        refreshSectionListModel();
        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
        refreshFloorListForSelectedSection();
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }


    // Основные операции с проектом
    private void loadProject(ActionEvent e) {
        try {
            List<Building> projects = DatabaseManager.getAllBuildings();
            if (projects.isEmpty()) {
                showMessage("Нет сохраненных проектов", "Информация", JOptionPane.INFORMATION_MESSAGE);
                return;
            }

            LoadProjectDialog dialog = new LoadProjectDialog(
                    (JFrame) SwingUtilities.getWindowAncestor(this), projects);
            dialog.setVisible(true);

            Building selectedProject = dialog.getSelectedProject();
            if (selectedProject != null) {
                loadSelectedProject(selectedProject);
            }
        } catch (SQLException ex) {
            handleError("Ошибка загрузки проектов: " + ex.getMessage(), "Ошибка");
        }
    }

    private void loadSelectedProject(Building selectedProject) throws SQLException {
        Building loadedBuilding = DatabaseManager.loadBuilding(selectedProject.getId());
        this.building = loadedBuilding;
        this.ops.setBuilding(this.building); // ← добавили
        projectNameField.setText(loadedBuilding.getName());
        refreshAllLists();
        updateVentilationTab(loadedBuilding);
        updateRadiationTab(loadedBuilding, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(loadedBuilding, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(loadedBuilding, /*autoApplyDefaults=*/false);
        showMessage("Проект '" + loadedBuilding.getName() + "' успешно загружен", "Загрузка", JOptionPane.INFORMATION_MESSAGE);
    }

    private void saveProject(ActionEvent e) {
        logger.info("BuildingTab.saveProject() - Начало сохранения проекта");

        // 0) Финализируем активное редактирование таблиц (чтобы текущее значение попало в модель)
        try {
            java.awt.KeyboardFocusManager kfm = java.awt.KeyboardFocusManager.getCurrentKeyboardFocusManager();
            java.awt.Component fo = (kfm != null) ? kfm.getFocusOwner() : null;
            JTable editingTable = (fo == null) ? null
                    : (JTable) javax.swing.SwingUtilities.getAncestorOfClass(JTable.class, fo);
            if (editingTable != null && editingTable.isEditing()) {
                try { editingTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        } catch (Exception ignore) {}

        // 1) Синхронизируем UI → модель (радиация)
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null) {
            radiationTab.updateRoomSelectionStates();
        }

// 1.1) КЕО (как и было)
        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

// чтобы именно эта вкладка окончательно зафиксировала is_selected
        ArtificialLightingTab artificialTab = getArtificialLightingTab();
        if (artificialTab != null) {
            artificialTab.updateRoomSelectionStates();
        }

// 1.3) Микроклимат (как и было)
        MicroclimateTab microTab = getMicroclimateTab();
        if (microTab != null) {
            microTab.updateRoomSelectionStates();
        }


        // 2) Генерация имени проекта
        String baseName = projectNameField.getText().trim();
        if (baseName.isEmpty()) {
            showMessage("Введите название проекта!", "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // 3) Создаем копию проекта с сохранением состояний
        Building newProject = createBuildingCopy();
        newProject.setName(generateProjectVersionName(baseName));

        // 4) Сохраняем в БД
        try {
            DatabaseManager.saveBuilding(newProject);
            this.building = newProject;
            this.ops.setBuilding(this.building); // ← добавили
            projectNameField.setText(extractBaseName(newProject.getName()));
        } catch (SQLException ex) {
            handleError("Ошибка сохранения: " + ex.getMessage(), "Ошибка");
            return;
        }
        refreshAllLists();

        // 5) Переинициализируем вкладки без авто-проставления
        updateRadiationTab(newProject, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(newProject, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(newProject, /*autoApplyDefaults=*/false);

        logger.info("Проект успешно сохранен");
    }

    private String generateProjectVersionName(String baseName) {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
        String currentDate = dateFormat.format(new Date());
        String cleanBaseName = extractBaseName(baseName);

        int nextVersion = calculateNextVersion(cleanBaseName);
        return nextVersion == 1 ?
                cleanBaseName :
                cleanBaseName + " ред." + nextVersion + " " + currentDate;
    }

    private int calculateNextVersion(String cleanBaseName) {
        try {
            List<Building> existingProjects = DatabaseManager.getAllBuildings();
            Pattern versionPattern = Pattern.compile("^" + Pattern.quote(cleanBaseName) + "(?: ред\\.(\\d+) (\\d{2}\\.\\d{2}\\.\\d{4}))?$");

            int maxVersion = existingProjects.stream()
                    .map(project -> versionPattern.matcher(project.getName()))
                    .filter(Matcher::find)
                    .mapToInt(matcher -> matcher.group(1) != null ? Integer.parseInt(matcher.group(1)) : 1)
                    .max()
                    .orElse(0);

            return maxVersion + 1;
        } catch (SQLException e) {
            handleError("Ошибка получения списка проектов: " + e.getMessage(), "Ошибка");
            return 1;
        }
    }

    private String extractBaseName(String name) {
        Pattern pattern = Pattern.compile("(.+?) (?:ред\\.\\d+ \\d{2}\\.\\d{2}\\.\\d{4})$");
        Matcher matcher = pattern.matcher(name);
        return matcher.find() ? matcher.group(1).trim() : name;
    }

    private Building createBuildingCopy() {
        Building copy = new Building();
        copy.setName(building.getName());

        // копируем секции
        List<Section> copiedSections = new ArrayList<>();
        for (Section s : building.getSections()) {
            Section cs = new Section();
            cs.setName(s.getName());
            cs.setPosition(s.getPosition());
            copiedSections.add(cs);
        }
        copy.setSections(copiedSections);

        // этажи — с сохранением selected и id
        for (Floor originalFloor : building.getFloors()) {
            Floor floorCopy = createFloorCopyPreserve(originalFloor);
            copy.addFloor(floorCopy);
        }
        return copy;
    }


    @SuppressWarnings("serial")
    private abstract class ReorderHandler<T> extends TransferHandler {
        private final DataFlavor flavor = DataFlavor.stringFlavor;
        @Override protected Transferable createTransferable(JComponent c) {
            @SuppressWarnings("unchecked") JList<T> list = (JList<T>) c;
            return new StringSelection(Integer.toString(list.getSelectedIndex()));
        }
        @Override public int getSourceActions(JComponent c) { return MOVE; }
        @Override public boolean canImport(TransferSupport s) { return s.isDrop() && s.isDataFlavorSupported(flavor); }
        @Override public boolean importData(TransferSupport s) {
            try {
                @SuppressWarnings("unchecked") JList<T> list = (JList<T>) s.getComponent();
                DefaultListModel<T> model = (DefaultListModel<T>) list.getModel();
                int to = ((JList.DropLocation) s.getDropLocation()).getIndex();
                int from = Integer.parseInt((String) s.getTransferable().getTransferData(flavor));
                if (from == -1 || from == to) return false;
                T elem = model.get(from);
                model.remove(from);
                if (to > from) to--;
                model.add(to, elem);
                list.setSelectedIndex(to);
                afterReorder(model);
                return true;
            } catch (Exception ex) { ex.printStackTrace(); return false; }
        }
        protected abstract void afterReorder(DefaultListModel<T> model);
    }

    private final TransferHandler spaceReorderHandler = new ReorderHandler<Space>() {
        @Override protected void afterReorder(DefaultListModel<Space> model) {
            // Переписываем position по новому порядку
            for (int i = 0; i < model.size(); i++) model.get(i).setPosition(i);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        }
    };

    private final TransferHandler roomReorderHandler = new ReorderHandler<Room>() {
        @Override protected void afterReorder(DefaultListModel<Room> model) {
            for (int i = 0; i < model.size(); i++) model.get(i).setPosition(i);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);


        }
    };

    private Floor createFloorCopy(Floor originalFloor) { return ops.createFloorCopy(originalFloor); }

    private void calculateMetrics(ActionEvent e) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            updateVentilationTab(building);
            updateRadiationTab(building, true);
            ; // Добавлено обновление RadiationTab
            ((MainFrame) mainFrame).selectVentilationTab();
        }
        showMessage("Данные для вентиляции обновлены!", "Расчет завершен", JOptionPane.INFORMATION_MESSAGE);
    }
    private void showApartmentSummary(ActionEvent e) {
        // Безопасность: если вдруг building == null, создаём пустой
        if (this.building == null) this.building = new Building();
        ApartmentSummaryDialog.show(this, this.building);
    }

    // Операции с помещениями
    private void copySpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        Space selectedSpace = spaceList.getSelectedValue();

        if (selectedFloor == null || selectedSpace == null) {
            showMessage("Выберите помещение для копирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Сохраняем состояния чекбоксов ДО копирования (радиация)
        Map<String, Boolean> savedSelections = saveRadiationSelections();

        // Создаём копию помещения с НОВЫМИ id комнат
        Space copiedSpace = createSpaceCopyWithNewIds(selectedSpace);
        copiedSpace.setIdentifier(generateNextSpaceId(selectedFloor, selectedSpace.getIdentifier()));

        // позиция в КОНЕЦ
        int maxPos = selectedFloor.getSpaces().stream()
                .mapToInt(Space::getPosition)
                .max().orElse(-1);
        copiedSpace.setPosition(maxPos + 1);


        // Добавляем
        selectedFloor.addSpace(copiedSpace);

        // Правильно пересобираем UI-список
        updateSpaceList();
        spaceList.setSelectedValue(copiedSpace, true);

        // Обновляем вкладки
        updateRadiationTab(building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        // Восстанавливаем состояния ТОЛЬКО для исходных комнат
        restoreRadiationSelections(savedSelections);

        // Явно выделим новое помещение в RadiationTab
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null) {
            int newIndex = selectedFloor.getSpaces().indexOf(copiedSpace);
            if (newIndex >= 0) radiationTab.selectSpaceByIndex(newIndex);
        }
    }

    private Space createSpaceCopyWithNewIds(Space originalSpace) { return ops.createSpaceCopyWithNewIds(originalSpace); }

    private String generateNextSpaceId(Floor floor, String currentId) {
        String prefix = "";
        String baseNumber = "";
        int lastDashIndex = currentId.lastIndexOf('-');

        if (lastDashIndex != -1) {
            prefix = currentId.substring(0, lastDashIndex).trim();
            baseNumber = currentId.substring(lastDashIndex + 1).trim();
        } else {
            prefix = currentId;
        }

        int currentNumber = parseNumber(baseNumber);
        int maxNumber = calculateMaxSpaceNumber(floor, prefix, currentNumber);
        int nextNumber = maxNumber + 1;

        return prefix.isEmpty() ?
                String.valueOf(nextNumber) :
                prefix + "-" + nextNumber;
    }

    private int parseNumber(String numberStr) {
        try {
            String numericPart = numberStr.replaceAll("\\D", "");
            return numericPart.isEmpty() ? 0 : Integer.parseInt(numericPart);
        } catch (NumberFormatException ignored) {
            return 0;
        }
    }

    private int calculateMaxSpaceNumber(Floor floor, String prefix, int currentNumber) {
        int maxNumber = currentNumber;
        for (Space space : floor.getSpaces()) {
            String id = space.getIdentifier();
            if (id.startsWith(prefix)) {
                String suffix = id.substring(prefix.length()).trim();
                if (suffix.startsWith("-")) suffix = suffix.substring(1).trim();

                int num = parseNumber(suffix);
                if (num > maxNumber) maxNumber = num;
            }
        }
        return maxNumber;
    }

    // Операции с этажами
    private void addFloor(ActionEvent e) {
        AddFloorDialog dialog = new AddFloorDialog((JFrame) SwingUtilities.getWindowAncestor(this));
        if (dialog.showDialog()) {
            Floor floor = new Floor();
            floor.setNumber(dialog.getFloorNumber());
            floor.setType(dialog.getFloorType());

            int secIdx = Math.max(0, sectionList.getSelectedIndex());
            floor.setSectionIndex(secIdx); // ← секция

            String floorName = floor.getType().title + " " + floor.getNumber();
            floor.setName(floorName);

// <<< НОВОЕ: ставим позицию в КОНЕЦ секции >>>
            int maxPos = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == secIdx)
                    .mapToInt(Floor::getPosition)
                    .max().orElse(-1);
            floor.setPosition(maxPos + 1);

            building.addFloor(floor);

            refreshFloorListForSelectedSection();
            updateRadiationTab(building, true);
            updateLightingTab(building, /*autoApplyDefaults=*/true);
            updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        }
    }

    private void updateRadiationTab(Building building, boolean forceOfficeSelection) {
        // ВАЖНО: по умолчанию авто-правила выключены, чтобы не сбивать ручные галочки
        updateRadiationTab(building, /*forceOfficeSelection=*/forceOfficeSelection, /*autoApplyRules=*/false);
    }

    private void updateRadiationTab(Building building, boolean forceOfficeSelection, boolean autoApplyRules) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                tab.setBuilding(building, forceOfficeSelection, autoApplyRules);
            }
        }
    }

    private RadiationTab getRadiationTab() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            return ((MainFrame) mainFrame).getRadiationTab();
        }
        return null;
    }

    private void copyFloor(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) return;

        // 1. Сохраняем ВСЕ текущие состояния комнат
        RadiationTab radiationTab = getRadiationTab();
        Map<Integer, Boolean> allRoomStates = new HashMap<>();
        if (radiationTab != null) {
            allRoomStates.putAll(radiationTab.globalRoomSelectionMap);
        }
        LightingTab lightingTab = getLightingTab();
        if (lightingTab != null) {
            lightingTab.updateRoomSelectionStates();
        }

        // 2. Создаем копию этажа
        Floor copiedFloor = createFloorCopy(selectedFloor);
        String newFloorNumber = generateNextFloorNumber(selectedFloor.getNumber(), selectedFloor.getSectionIndex());
        copiedFloor.setNumber(newFloorNumber);
        copiedFloor.setSectionIndex(selectedFloor.getSectionIndex());

// ВАЖНО: позицию ставим в КОНЕЦ секции → копия не первый этаж
        int secIdx = selectedFloor.getSectionIndex();
        int maxPos = building.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .mapToInt(Floor::getPosition)
                .max().orElse(-1);
        copiedFloor.setPosition(maxPos + 1);

        updateSpaceIdentifiers(copiedFloor, extractDigits(newFloorNumber));

        // 3. Добавляем новый этаж в модель
        building.addFloor(copiedFloor);
        floorListModel.addElement(copiedFloor);

        // 4. Обновляем ТОЛЬКО список этажей в UI
        floorList.setSelectedValue(copiedFloor, true);
// Освещение: при копировании этажа авто-правила НЕ применяем — галочки должны быть пустыми
        updateLightingTab(building, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        // 5. Восстанавливаем ВСЕ состояния комнат
        if (radiationTab != null) {
            radiationTab.globalRoomSelectionMap.clear();
            radiationTab.globalRoomSelectionMap.putAll(allRoomStates);

            // 6. Устанавливаем состояния для новых комнат
            for (int i = 0; i < selectedFloor.getSpaces().size(); i++) {
                Space origSpace = selectedFloor.getSpaces().get(i);
                Space copiedSpace = copiedFloor.getSpaces().get(i);

                for (int j = 0; j < origSpace.getRooms().size(); j++) {
                    Room origRoom = origSpace.getRooms().get(j);
                    Room copiedRoom = copiedSpace.getRooms().get(j);

                    Boolean state = allRoomStates.get(origRoom.getId());
                    if (state != null) {
                        radiationTab.setRoomSelectionState(copiedRoom.getId(), state);
                    }
                }
            }

            // 7. Обновляем UI RadiationTab
            radiationTab.refreshFloors();
        }
    }

    private String extractDigits(String input) { return ops.extractDigits(input); }

    // Новый: считает следующий номер ТОЛЬКО в пределах указанной секции.
    private String generateNextFloorNumber(String currentNumber, int sectionIndex) {
        return ops.generateNextFloorNumber(currentNumber, sectionIndex);
    }

    private void updateSpaceIdentifiers(Floor floor, String newFloorDigits) { ops.updateSpaceIdentifiers(floor, newFloorDigits); }


    private String updateIdentifier(String identifier, String newFloorDigits) { return ops.updateIdentifier(identifier, newFloorDigits); }

    private void editFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        if (index < 0) {
            showMessage("Выберите этаж для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Floor floor = floorListModel.get(index);
        AddFloorDialog dialog = new AddFloorDialog((JFrame) SwingUtilities.getWindowAncestor(this));
        dialog.setFloorNumber(floor.getNumber());
        dialog.setFloorType(floor.getType());

        if (dialog.showDialog()) {
            floor.setNumber(dialog.getFloorNumber());
            floor.setType(dialog.getFloorType());
            floorListModel.set(index, floor);
            updateSpaceList();

        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }

    private Map<String, Boolean> saveRadiationSelections() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                return tab.saveSelections();
            }
        }
        return new HashMap<>();
    }

    private void restoreRadiationSelections(Map<String, Boolean> savedSelections) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            RadiationTab tab = ((MainFrame) mainFrame).getRadiationTab();
            if (tab != null) {
                tab.restoreSelections(savedSelections);
            }
        }
    }

    private void removeFloor(ActionEvent e) {
        int index = floorList.getSelectedIndex();
        if (index >= 0) {
            building.getFloors().remove(index);
            floorListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }

    // Операции с помещениями
    private void addSpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) {
            showMessage("Выберите этаж!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), selectedFloor.getType());
        Space space = null;
        if (dialog.showDialog()) {
            space = new Space();
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());

            // позиция в КОНЕЦ текущего этажа
            int maxPos = selectedFloor.getSpaces().stream()
                    .mapToInt(Space::getPosition)
                    .max().orElse(-1);
            space.setPosition(maxPos + 1);

            selectedFloor.addSpace(space);

            // офисам — МК по умолчанию (кроме санузлов)
            applyMicroDefaultsForOfficeSpace(space);

            // ПЕРЕСОБРАТЬ список с учётом сортировки и фильтров
            updateSpaceList();
            if (spaceListModel.getSize() > 0) {
                spaceList.setSelectedValue(space, true);
            }
        }

        // обновляем вкладки
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

        // фокус в RadiationTab на новое помещение
        RadiationTab radiationTab = getRadiationTab();
        if (radiationTab != null && space != null) {
            Floor f = floorList.getSelectedValue();
            if (f != null) {
                int newIndex = f.getSpaces().indexOf(space);
                if (newIndex >= 0) radiationTab.selectSpaceByIndex(newIndex);
            }
        }
    }

    private void editSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();

        if (floor == null || index < 0) {
            showMessage("Выберите помещение для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Space space = spaceListModel.get(index);
        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), floor.getType());
        dialog.setSpaceIdentifier(space.getIdentifier());
        dialog.setSpaceType(space.getType());

        if (dialog.showDialog()) {
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());
            applyMicroDefaultsForOfficeSpace(space);
            spaceListModel.set(index, space);
            updateRoomList();
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);


    }

    private void removeSpace(ActionEvent e) {
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();
        if (floor != null && index >= 0) {
            floor.getSpaces().remove(index);
            spaceListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }


    // «умное» добавление комнат — с подсказками, пакетным добавлением и автонумерацией
    private void addRoom(ActionEvent e) {
        Space selectedSpace = spaceList.getSelectedValue();
        if (selectedSpace == null) {
            showMessage("Выберите помещение!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (isApartmentSpace(selectedSpace)) {
            // квартиры — AddRoomDialog
            java.util.List<String> suggestions = collectPopularApartmentRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddRoomDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddRoomDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                java.util.List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else if (isOfficeSpace(selectedSpace)) {
            // офисы — AddOfficeRoomsDialog
            java.util.List<String> suggestions = collectPopularOfficeRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                java.util.List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    // микроклимат в офисе — автоматом, кроме санузлов
                    room.setMicroclimateSelected(!looksLikeSanitary(room.getName()));
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else if (isPublicSpace(selectedSpace)) {
            // ОБЩЕСТВЕННЫЕ — AddPublicRoomsDialog (никаких офис/жилых автопроставлений)
            java.util.List<String> suggestions = collectPopularPublicRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog(parent, suggestions, "", false);
            if (dlg.showDialog()) {
                java.util.List<String> names = dlg.getNamesToAddList();
                for (String name : names) {
                    if (name == null || name.isBlank()) continue;
                    Room room = new Room();
                    room.setName(name.trim());
                    room.setSelected(false);
                    // микроклимат — БЕЗ автопроставления (по твоему требованию)
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    selectedSpace.addRoom(room);
                    roomListModel.addElement(room);
                }
            }
        } else {
            // прочие типы — простое поле ввода
            String name = showInputDialog("Добавление комнаты", "Название комнаты:", "");
            if (name != null && !name.isBlank()) {
                Room room = new Room();
                room.setName(name.trim());
                room.setSelected(false);
                try {
                    int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                    room.setExternalWallsCount(walls);
                } catch (Throwable ignore) {}
                selectedSpace.addRoom(room);
                roomListModel.addElement(room);
            }
        }

        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }


    private void editRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space == null || index < 0) {
            showMessage("Выберите комнату для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Room room = roomListModel.get(index);

        if (isApartmentSpace(space)) {
            java.util.List<String> suggestions = collectPopularApartmentRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddRoomDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddRoomDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    roomListModel.set(index, room);
                    updateRadiationTab(building, true);
                    updateLightingTab(building, /*autoApplyDefaults=*/true);
                    updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
                }
            }
        } else if (isOfficeSpace(space)) {
            java.util.List<String> suggestions = collectPopularOfficeRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddOfficeRoomsDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    // офис: подсвежим дефолты
                    room.setMicroclimateSelected(!looksLikeSanitary(room.getName()));
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    roomListModel.set(index, room);
                    updateRadiationTab(building, true);
                    updateLightingTab(building, /*autoApplyDefaults=*/true);
                    updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
                }
            }
        } else if (isPublicSpace(space)) {
            java.util.List<String> suggestions = collectPopularPublicRoomNames(building, 30);
            JFrame parent = (JFrame) SwingUtilities.getWindowAncestor(this);
            ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog dlg =
                    new ru.citlab24.protokol.tabs.dialogs.AddPublicRoomsDialog(parent, suggestions, room.getName(), true);
            if (dlg.showDialog()) {
                String newName = dlg.getNameToAdd();
                if (newName != null && !newName.isBlank()) {
                    room.setName(newName.trim());
                    // для общественных автопроставления МК по умолчанию не делаем
                    try {
                        int walls = looksLikeSanitary(room.getName()) ? 0 : 1;
                        room.setExternalWallsCount(walls);
                    } catch (Throwable ignore) {}
                    roomListModel.set(index, room);
                    updateRadiationTab(building, true);
                    updateLightingTab(building, /*autoApplyDefaults=*/true);
                    updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
                }
            }
        } else {
            String newName = showInputDialog("Редактирование комнаты", "Новое название комнаты:", room.getName());
            if (newName != null && !newName.isBlank()) {
                room.setName(newName.trim());
                roomListModel.set(index, room);
                updateRadiationTab(building, true);
                updateLightingTab(building, /*autoApplyDefaults=*/true);
                updateMicroclimateTab(building, /*autoApplyDefaults=*/false);
            }
        }

    }


    private void removeRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space != null && index >= 0) {
            space.getRooms().remove(index);
            roomListModel.remove(index);
        }
        updateRadiationTab(building, true);
        updateLightingTab(building, /*autoApplyDefaults=*/true);
        updateMicroclimateTab(building, /*autoApplyDefaults=*/false);

    }

    private void updateSpaceList() {
        spaceListModel.clear();
        Floor selectedFloor = (floorList != null) ? floorList.getSelectedValue() : null;
        if (selectedFloor == null) return;

        java.util.List<Space> sorted = new java.util.ArrayList<>(selectedFloor.getSpaces());
        sorted.sort(java.util.Comparator.comparingInt(Space::getPosition));

        for (Space s : sorted) {
            if (isSpaceVisibleByFilter(s)) spaceListModel.addElement(s);
        }

        if (!spaceListModel.isEmpty()) {
            Space prev = spaceList.getSelectedValue();
            int idx = (prev != null) ? spaceListModel.indexOf(prev) : -1;
            spaceList.setSelectedIndex(idx >= 0 ? idx : 0);
        } else {
            roomListModel.clear();
        }
    }

    private void updateRoomList() {
        roomListModel.clear();
        Space selectedSpace = spaceList.getSelectedValue();

        if (selectedSpace != null) {
            for (Room room : selectedSpace.getRooms()) {
                roomListModel.addElement(room);
            }
        }
    }

    private void refreshAllLists() {
        // Секции
        refreshSectionListModel();
        if (!sectionListModel.isEmpty() && sectionList.getSelectedIndex() < 0) {
            sectionList.setSelectedIndex(0);
        }

        // Этажи выбранной секции
        refreshFloorListForSelectedSection();

        // Помещения/комнаты для актуального выбора
        updateSpaceList();
        updateRoomList();
    }

    private void updateLightingTab(Building building, boolean autoApplyDefaults) {
        // КЕО — как было
        LightingTab keo = getLightingTab();
        if (keo != null) {
            keo.setBuilding(building, autoApplyDefaults);
            keo.refreshData();
        }
        // НОВОЕ: отдельная вкладка «Освещение» со своей картой галочек
        ArtificialLightingTab alt = getArtificialLightingTab();
        if (alt != null) {
            alt.setBuilding(building, autoApplyDefaults); // при true – авто-галочки для офис/общественных
            alt.refreshData();
        }
    }

    private void updateVentilationTab(Building building) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            Arrays.stream(((MainFrame) mainFrame).getTabbedPane().getComponents())
                    .filter(tab -> tab instanceof VentilationTab)
                    .findFirst()
                    .ifPresent(tab -> {
                        ((VentilationTab) tab).setBuilding(building);
                        ((VentilationTab) tab).refreshData();
                    });
        }
    }

    private String showInputDialog(String title, String message, String initialValue) {
        JTextField field = new JTextField(initialValue);

        JPanel panel = new JPanel(new BorderLayout(8, 0));
        panel.add(new JLabel(message), BorderLayout.WEST);
        panel.add(field, BorderLayout.CENTER);

        // Создаём диалог вручную, чтобы управлять фокусом и клавишами
        final JOptionPane pane = new JOptionPane(
                panel,
                JOptionPane.PLAIN_MESSAGE,
                JOptionPane.OK_CANCEL_OPTION
        );
        final JDialog dialog = pane.createDialog(this, title);
        dialog.setModal(true);
        dialog.setResizable(false);

        // Enter = OK
        field.addActionListener(e -> {
            pane.setValue(JOptionPane.OK_OPTION);
            dialog.dispose();
        });

        // Esc = Cancel
        dialog.getRootPane().registerKeyboardAction(
                e -> {
                    pane.setValue(JOptionPane.CANCEL_OPTION);
                    dialog.dispose();
                },
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW
        );

        // Автофокус и выделение текста сразу после показа
        dialog.addWindowListener(new WindowAdapter() {
            @Override public void windowOpened(WindowEvent e) {
                field.requestFocusInWindow();
                field.selectAll();
            }
        });

        dialog.setVisible(true);

        Object value = pane.getValue();
        boolean ok = (value != null) && Integer.valueOf(JOptionPane.OK_OPTION).equals(value);
        return ok ? field.getText() : null;
    }

    private void showMessage(String message, String title, int messageType) {
        JOptionPane.showMessageDialog(this, message, title, messageType);
    }

    private void handleError(String message, String title) {
        // Логирование ошибки
        System.err.println(title + ": " + message);
        showMessage(message, title, JOptionPane.ERROR_MESSAGE);
    }

    @Override
    public void addNotify() {
        super.addNotify();
        if (this.building == null) this.building = new Building();
        updateRadiationTab(this.building, /*forceOfficeSelection=*/false, /*autoApplyRules=*/false);
        updateLightingTab(building, /*autoApplyDefaults=*/false);
        updateMicroclimateTab(building, false);

        if (floorList != null && !floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        }
        if (floorList != null) {
            floorList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) updateSpaceList();
            });
        }
        if (spaceList != null) {
            spaceList.addListSelectionListener(evt -> {
                if (!evt.getValueIsAdjusting()) updateRoomList();
            });
        }
    }

    private java.util.List<String> collectPopularApartmentRoomNames(Building b, int limit) {
        return ops.collectPopularApartmentRoomNames(limit);
    }

    // Возвращает true, если помещение — «квартира».
// Не опираемся на конкретные enum-константы проекта, работаем по имени типа и по идентификатору.
    private boolean isApartmentSpace(Space s) { return ops.isApartmentSpace(s); }


    // Офисное помещение?
    private boolean isOfficeSpace(Space s) { return ops.isOfficeSpace(s); }

    // Общественное помещение?
    private boolean isPublicSpace(Space s) { return ops.isPublicSpace(s); }

    // Топ общественных названий по зданию
    private java.util.List<String> collectPopularPublicRoomNames(Building b, int limit) {
        return ops.collectPopularPublicRoomNames(limit);
    }

    private static String normalizeOfficeBaseName(String name) {
        return BuildingModelOps.normalizeOfficeBaseName(name);
    }

    // Сбор топа офисных названий по зданию (по "базе", без номеров)
    private java.util.List<String> collectPopularOfficeRoomNames(Building b, int limit) {
        return ops.collectPopularOfficeRoomNames(limit);
    }

    // КОПИЯ ДЛЯ СОХРАНЕНИЯ: сохраняем id и selected
    private Floor createFloorCopyPreserve(Floor original) { return ops.createFloorCopyPreserve(original); }


    private Space createSpaceCopyPreserve(Space original) { return ops.createSpaceCopyPreserve(original); }

    private LightingTab getLightingTab() {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof MainFrame) {
            JTabbedPane tabs = ((MainFrame) wnd).getTabbedPane();
            for (Component c : tabs.getComponents()) {
                if (c instanceof LightingTab) return (LightingTab) c;
            }
        }
        return null;
    }
    private ArtificialLightingTab getArtificialLightingTab() {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof MainFrame) {
            JTabbedPane tabs = ((MainFrame) wnd).getTabbedPane();
            for (Component c : tabs.getComponents()) {
                if (c instanceof ArtificialLightingTab) return (ArtificialLightingTab) c;
            }
        }
        return null;
    }

    private void updateMicroclimateTab(Building building, boolean autoApplyDefaults) {
        Window wnd = SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof MainFrame) {
            JTabbedPane tabs = ((MainFrame) wnd).getTabbedPane();
            for (Component c : tabs.getComponents()) {
                if (c instanceof MicroclimateTab) {
                    ((MicroclimateTab) c).display(building, autoApplyDefaults);
                    break;
                }
            }
        }
    }
    private MicroclimateTab getMicroclimateTab() {
        java.awt.Window wnd = javax.swing.SwingUtilities.getWindowAncestor(this);
        if (wnd instanceof ru.citlab24.protokol.MainFrame) {
            javax.swing.JTabbedPane tabs = ((ru.citlab24.protokol.MainFrame) wnd).getTabbedPane();
            for (java.awt.Component c : tabs.getComponents()) {
                if (c instanceof ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab) {
                    return (ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateTab) c;
                }
            }
        }
        return null;
    }
    private static boolean looksLikeSanitary(String name) {
        return BuildingModelOps.looksLikeSanitary(name);
    }

    // ===== МИКРОКЛИМАТ: сохранение/восстановление по ключу "этаж|помещение|комната" =====
    private Map<String, Boolean> saveMicroclimateSelections() { return ops.saveMicroclimateSelections(); }

    private void restoreMicroclimateSelections(Map<String, Boolean> saved) { ops.restoreMicroclimateSelections(saved); }

    /** Микроклимат: для офисов (по типу помещения ИЛИ по типу этажа) ставим всем, кроме санузлов */
    private void applyMicroDefaultsForOfficeSpace(Space s) { ops.applyMicroDefaultsForOfficeSpace(s); }

    /** true, если помещение проходит текущие тумблеры-фильтры. */
    private boolean isSpaceVisibleByFilter(Space s) {
        boolean showA = (filterApartmentBtn == null) || filterApartmentBtn.isSelected();
        boolean showO = (filterOfficeBtn    == null) || filterOfficeBtn.isSelected();
        boolean showP = (filterPublicBtn    == null) || filterPublicBtn.isSelected();

        boolean isA = isApartmentSpace(s);
        boolean isO = isOfficeSpace(s);
        boolean isP = isPublicSpace(s);

        if (isA && showA) return true;
        if (isO && showO) return true;
        if (isP && showP) return true;

        // Нераспознанный тип показываем только если включены все тумблеры
        if (!isA && !isO && !isP) return showA && showO && showP;

        return false;
    }

}