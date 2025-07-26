package ru.citlab24.protokol.tabs.buildingTab;

import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.modules.med.RadiationTab;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationTab;
import ru.citlab24.protokol.tabs.dialogs.AddFloorDialog;
import ru.citlab24.protokol.tabs.dialogs.AddSpaceDialog;
import ru.citlab24.protokol.tabs.dialogs.LoadProjectDialog;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.RoomListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.function.Supplier;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class BuildingTab extends JPanel {
    // Константы для цветов и стилей
    private static final Color FLOOR_PANEL_COLOR = new Color(0, 115, 200);
    private static final Color SPACE_PANEL_COLOR = new Color(76, 175, 80);
    private static final Color ROOM_PANEL_COLOR = new Color(156, 39, 176);
    private static final Font HEADER_FONT = new Font("Segoe UI", Font.BOLD, 14);
    private static final Dimension BUTTON_PANEL_SIZE = new Dimension(5, 5);

    private Building building;
    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final DefaultListModel<Space> spaceListModel = new DefaultListModel<>();
    private final DefaultListModel<Room> roomListModel = new DefaultListModel<>();

    private JList<Floor> floorList;
    private JList<Space> spaceList;
    private JList<Room> roomList;
    private JTextField projectNameField;

    public BuildingTab(Building building) {
        setRussianLocale();
        this.building = building;
        initComponents();
    }

    private void setRussianLocale() {
        Locale.setDefault(new Locale("ru", "RU"));
        UIManager.put("OptionPane.yesButtonText", "Да");
        UIManager.put("OptionPane.noButtonText", "Нет");
        UIManager.put("OptionPane.cancelButtonText", "Отмена");

        // Оптимизированная установка локали для UIManager
        UIManager.getLookAndFeelDefaults().keySet().stream()
                .filter(key -> key instanceof String && ((String) key).endsWith(".locale"))
                .forEach(key -> UIManager.getLookAndFeelDefaults().put(key, Locale.getDefault()));
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
        projectNameField.setText(building.getName());

        panel.add(label);
        panel.add(projectNameField);

        return panel;
    }

    private JPanel createBuildingPanel() {
        JPanel panel = new JPanel(new GridLayout(1, 3, 15, 15));
        panel.add(createListPanel("Этажи здания", FLOOR_PANEL_COLOR, floorListModel,
                new FloorListRenderer(), this::createFloorButtons));
        panel.add(createListPanel("Помещения на этаже", SPACE_PANEL_COLOR, spaceListModel,
                new SpaceListRenderer(), this::createSpaceButtons));
        panel.add(createListPanel("Комнаты в помещении", ROOM_PANEL_COLOR, roomListModel,
                new RoomListRenderer(), this::createRoomButtons));
        return panel;
    }

    private <T> JPanel createListPanel(String title, Color color, ListModel<T> model,
                                       ListCellRenderer<? super T> renderer,
                                       Supplier<JPanel> buttonPanelSupplier) {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder(title, color));

        JList<T> list = new JList<>(model);
        list.setCellRenderer(renderer);
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Сохраняем ссылки на списки
        if ("Этажи здания".equals(title)) {
            floorList = (JList<Floor>) list;
        } else if ("Помещения на этаже".equals(title)) {
            spaceList = (JList<Space>) list;
        } else if ("Комнаты в помещении".equals(title)) {
            roomList = (JList<Room>) list;
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
                createStyledButton("", FontAwesomeSolid.PLUS, new Color(46, 125, 50), this::addFloor),
                createStyledButton("", FontAwesomeSolid.CLONE, new Color(100, 181, 246), this::copyFloor),
                createStyledButton("", FontAwesomeSolid.EDIT, new Color(255, 152, 0), this::editFloor),
                createStyledButton("", FontAwesomeSolid.TRASH, new Color(198, 40, 40), this::removeFloor)
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
        return createButtonPanel(
                createStyledButton("Загрузить проект", FontAwesomeSolid.FOLDER_OPEN, new Color(33, 150, 243), this::loadProject),
                createStyledButton("Сохранить проект", FontAwesomeSolid.SAVE, new Color(0, 115, 200), this::saveProject),
                createStyledButton("Рассчитать показатели", FontAwesomeSolid.CALCULATOR, new Color(103, 58, 183), this::calculateMetrics)
        );
    }

    private JButton createStyledButton(String text, FontAwesomeSolid icon, Color bgColor, ActionListener action) {
        JButton btn = new JButton(text, FontIcon.of(icon, 16, Color.WHITE));
        btn.setBackground(bgColor);
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.setFont(HEADER_FONT);
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
        projectNameField.setText(loadedBuilding.getName());
        refreshAllLists();
        updateVentilationTab(loadedBuilding);
        updateRadiationTab(loadedBuilding); // Добавлено обновление RadiationTab
        showMessage("Проект '" + loadedBuilding.getName() + "' успешно загружен", "Загрузка", JOptionPane.INFORMATION_MESSAGE);
    }

    private void saveProject(ActionEvent e) {
        try {
            // Сохраняем состояния чекбоксов ДО создания копии
            Map<String, Boolean> radiationSelections = saveRadiationSelections();

            String baseName = projectNameField.getText().trim();
            if (baseName.isEmpty()) {
                showMessage("Введите название проекта!", "Ошибка", JOptionPane.ERROR_MESSAGE);
                return;
            }

            String finalName = generateProjectVersionName(baseName);
            saveVentilationCalculations();

            Building newProject = createBuildingCopy();
            newProject.setName(finalName);
            DatabaseManager.saveBuilding(newProject);

            this.building = newProject;
            projectNameField.setText(extractBaseName(finalName));

            // Обновляем вкладку с новым зданием
            updateRadiationTab(newProject);

            // Восстанавливаем состояния ПОСЛЕ обновления
            restoreRadiationSelections(radiationSelections);

            showMessage("Новая версия проекта сохранена как: " + finalName, "Сохранение", JOptionPane.INFORMATION_MESSAGE);
            updateVentilationTab(newProject);
        } catch (SQLException ex) {
            handleError("Ошибка сохранения проекта: " + ex.getMessage(), "Ошибка");
        }
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
                    .orElse(1);

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

        for (Floor originalFloor : building.getFloors()) {
            Floor floorCopy = createFloorCopy(originalFloor);
            copy.addFloor(floorCopy);
        }
        return copy;
    }

    private Floor createFloorCopy(Floor originalFloor) {
        Floor floorCopy = new Floor();
        floorCopy.setNumber(originalFloor.getNumber());
        floorCopy.setType(originalFloor.getType());
        floorCopy.setName(originalFloor.getName()); // Копируем имя если нужно

        // Глубокое копирование помещений
        for (Space originalSpace : originalFloor.getSpaces()) {
            Space spaceCopy = new Space();
            spaceCopy.setIdentifier(originalSpace.getIdentifier());
            spaceCopy.setType(originalSpace.getType());

            // Глубокое копирование комнат
            for (Room originalRoom : originalSpace.getRooms()) {
                Room roomCopy = new Room();
                roomCopy.setId(originalRoom.getId());
                roomCopy.setName(originalRoom.getName());
                roomCopy.setVolume(originalRoom.getVolume());
                roomCopy.setVentilationChannels(originalRoom.getVentilationChannels());
                roomCopy.setVentilationSectionArea(originalRoom.getVentilationSectionArea());
                spaceCopy.addRoom(roomCopy);
            }
            floorCopy.addSpace(spaceCopy);
        }
        return floorCopy;
    }

    private Space createSpaceCopy(Space originalSpace) {
        Space spaceCopy = new Space();
        spaceCopy.setIdentifier(originalSpace.getIdentifier());
        spaceCopy.setType(originalSpace.getType());

        for (Room originalRoom : originalSpace.getRooms()) {
            Room roomCopy = createRoomCopy(originalRoom);
            spaceCopy.addRoom(roomCopy);
        }
        return spaceCopy;
    }

    private Room createRoomCopy(Room originalRoom) {
        Room roomCopy = new Room();
        roomCopy.setId(originalRoom.getId()); // Сохраняем оригинальный ID!
        roomCopy.setName(originalRoom.getName());
        roomCopy.setVolume(originalRoom.getVolume());
        roomCopy.setVentilationChannels(originalRoom.getVentilationChannels());
        roomCopy.setVentilationSectionArea(originalRoom.getVentilationSectionArea());
        return roomCopy;
    }

    private int generateUniqueRoomId() {
        return UUID.randomUUID().hashCode();
    }

    private void calculateMetrics(ActionEvent e) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            updateVentilationTab(building);
            updateRadiationTab(building); // Добавлено обновление RadiationTab
            ((MainFrame) mainFrame).selectVentilationTab();
        }
        showMessage("Данные для вентиляции обновлены!", "Расчет завершен", JOptionPane.INFORMATION_MESSAGE);
    }

    // Операции с помещениями
    private void copySpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        Space selectedSpace = spaceList.getSelectedValue();

        if (selectedFloor == null || selectedSpace == null) {
            showMessage("Выберите помещение для копирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Space copiedSpace = createSpaceCopy(selectedSpace);
        copiedSpace.setIdentifier(generateNextSpaceId(selectedFloor, selectedSpace.getIdentifier()));

        selectedFloor.addSpace(copiedSpace);
        spaceListModel.addElement(copiedSpace);
        spaceList.setSelectedValue(copiedSpace, true);
        updateRadiationTab(building);
    }

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

            // Устанавливаем имя этажа как комбинацию типа и номера
            String floorName = floor.getType().title + " " + floor.getNumber();
            floor.setName(floorName);

            building.addFloor(floor);
            floorListModel.addElement(floor);
            updateRadiationTab(building);
        }
    }
    private void updateRadiationTab(Building building) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            MainFrame frame = (MainFrame) mainFrame;
            RadiationTab tab = frame.getRadiationTab(); // Нужно добавить метод доступа

            if (tab != null) {
                tab.setBuilding(building);
//                tab.refreshData(); // Добавить этот метод в RadiationTab
            }
        }
    }
    private void copyFloor(ActionEvent e) {
        Map<String, Boolean> savedSelections = saveRadiationSelections();
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) return;

        Floor copiedFloor = createFloorCopy(selectedFloor);
        String newFloorNumber = generateNextFloorNumber(selectedFloor.getNumber());
        copiedFloor.setNumber(newFloorNumber);

        updateSpaceIdentifiers(copiedFloor, extractDigits(newFloorNumber));
        building.addFloor(copiedFloor);
        floorListModel.addElement(copiedFloor);
        floorList.setSelectedValue(copiedFloor, true);

        // Восстанавливаем ПОСЛЕ добавления этажа
        restoreRadiationSelections(savedSelections);
        updateRadiationTab(building);
    }
    private String extractDigits(String input) {
        return input.replaceAll("\\D", ""); // Удаляем все не-цифры
    }
    private String generateNextFloorNumber(String currentNumber) {
        // Пытаемся извлечь число из названия этажа
        Pattern p = Pattern.compile("\\d+");
        Matcher m = p.matcher(currentNumber);

        if (m.find()) {
            int num = Integer.parseInt(m.group());
            String prefix = currentNumber.substring(0, m.start());
            String suffix = currentNumber.substring(m.end());

            // Ищем максимальный существующий номер
            int maxNumber = building.getFloors().stream()
                    .map(Floor::getNumber)
                    .mapToInt(n -> {
                        Matcher matcher = p.matcher(n);
                        return matcher.find() ? Integer.parseInt(matcher.group()) : Integer.MIN_VALUE;
                    })
                    .max()
                    .orElse(num);

            // Генерируем следующий номер
            int nextNumber = (maxNumber != Integer.MIN_VALUE) ? maxNumber + 1 : num + 1;
            return prefix + nextNumber + suffix;
        }

        // Для нечисловых названий
        return generateUniqueNonNumericName(currentNumber);
    }
    private String findNextAvailableNumber(int baseNumber) {
        int candidate = baseNumber + 1;
        Set<Integer> existingNumbers = new HashSet<>();

        // Собираем все числовые номера этажей
        for (Floor f : building.getFloors()) {
            try {
                existingNumbers.add(Integer.parseInt(f.getNumber()));
            } catch (NumberFormatException ignored) {
                // Игнорируем нечисловые этажи
            }
        }

        // Ищем ближайшее свободное число
        while (existingNumbers.contains(candidate)) {
            candidate++;
        }
        return String.valueOf(candidate);
    }

    private String generateUniqueNonNumericName(String base) {
        Pattern pattern = Pattern.compile(Pattern.quote(base) + "(?: \\(копия (\\d+)\\))?");
        int maxCopy = building.getFloors().stream()
                .map(Floor::getNumber)
                .map(pattern::matcher)
                .filter(Matcher::matches)
                .mapToInt(m -> m.group(1) != null ? Integer.parseInt(m.group(1)) : 0)
                .max()
                .orElse(0);

        return base + (maxCopy == 0 ? "" : " (копия " + (maxCopy + 1) + ")");
    }
    private void updateSpaceIdentifiers(Floor floor, String newFloorDigits) {
        for (Space space : floor.getSpaces()) {
            String newId = updateIdentifier(space.getIdentifier(), newFloorDigits);
            space.setIdentifier(newId);
        }
    }
    private String updateIdentifier(String identifier, String newFloorDigits) {
        // Шаблон для поиска формата "X-Y" (например: "кв 1-1")
        Pattern pattern = Pattern.compile("(.*?)(\\d+)-(\\d+)(.*)");
        Matcher matcher = pattern.matcher(identifier);

        if (matcher.matches()) {
            String prefix = matcher.group(1);  // Префикс (например, "кв ")
            String roomNum = matcher.group(3);  // Номер помещения
            String suffix = matcher.group(4);   // Суффикс (если есть)

            // Формируем новый идентификатор: префикс + цифры этажа + номер помещения
            return prefix + newFloorDigits + "-" + roomNum + suffix;
        }
        return identifier; // Возвращаем оригинал, если паттерн не совпал
    }

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
            // При смене типа на PUBLIC - создать помещение при необходимости
            if (floor.getType() == Floor.FloorType.PUBLIC) {
                createDefaultSpaceIfMissing(floor);
            }
            floorListModel.set(index, floor);
            updateSpaceList();
        }
        updateRadiationTab(building);
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
        updateRadiationTab(building);
    }

    // Операции с помещениями
    private void addSpace(ActionEvent e) {
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) {
            showMessage("Выберите этаж!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        AddSpaceDialog dialog = new AddSpaceDialog((JFrame) SwingUtilities.getWindowAncestor(this), selectedFloor.getType());
        if (dialog.showDialog()) {
            Space space = new Space();
            space.setIdentifier(dialog.getSpaceIdentifier());
            space.setType(dialog.getSpaceType());
            selectedFloor.addSpace(space);
            spaceListModel.addElement(space);
        }
        updateRadiationTab(building);
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
            spaceListModel.set(index, space);
            updateRoomList();
        }
        updateRadiationTab(building);
    }

    private void removeSpace(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        if (space != null && "-".equals(space.getIdentifier())) {
            showMessage("Нельзя удалить системное помещение", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }
        Floor floor = floorList.getSelectedValue();
        int index = spaceList.getSelectedIndex();

        if (floor != null && index >= 0) {
            floor.getSpaces().remove(index);
            spaceListModel.remove(index);
        }
        updateRadiationTab(building);
    }

    // Операции с комнатами
    private void addRoom(ActionEvent e) {
        Space selectedSpace = spaceList.getSelectedValue();
        if (selectedSpace == null) {
            showMessage("Выберите помещение!", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        String name = showInputDialog("Добавление комнаты", "Название комнаты:", "");
        if (name != null && !name.trim().isEmpty()) {
            Room room = new Room();
            room.setName(name.trim());
            selectedSpace.addRoom(room);
            roomListModel.addElement(room);
        }
        updateRadiationTab(building);
    }

    private void editRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space == null || index < 0) {
            showMessage("Выберите комнату для редактирования", "Ошибка", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Room room = roomListModel.get(index);
        String newName = showInputDialog("Редактирование комнаты", "Новое название комнаты:", room.getName());
        if (newName != null && !newName.trim().isEmpty()) {
            room.setName(newName.trim());
            roomListModel.set(index, room);
        }
        updateRadiationTab(building);
    }

    private void removeRoom(ActionEvent e) {
        Space space = spaceList.getSelectedValue();
        int index = roomList.getSelectedIndex();

        if (space != null && index >= 0) {
            space.getRooms().remove(index);
            roomListModel.remove(index);
        }
        updateRadiationTab(building);
    }

    // Вспомогательные методы
    private void updateSpaceList() {
        spaceListModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();

        if (selectedFloor != null) {
            // Для общественных этажей: проверяем и создаем помещение "-" при необходимости
            if (selectedFloor.getType() == Floor.FloorType.PUBLIC) {
                createDefaultSpaceIfMissing(selectedFloor);
            }

            // Заполняем список помещений
            selectedFloor.getSpaces().forEach(spaceListModel::addElement);

            // Автоматически выбираем помещение "-" для общественных этажей
            if (selectedFloor.getType() == Floor.FloorType.PUBLIC && !spaceListModel.isEmpty()) {
                selectDefaultSpace(selectedFloor);
            } else if (!spaceListModel.isEmpty()) {
                spaceList.setSelectedIndex(0);
            }
        }
    }
    private void createDefaultSpaceIfMissing(Floor floor) {
        boolean hasDefaultSpace = floor.getSpaces().stream()
                .anyMatch(space -> "-".equals(space.getIdentifier()));

        if (!hasDefaultSpace) {
            Space defaultSpace = new Space();
            defaultSpace.setIdentifier("-");
            defaultSpace.setType(Space.SpaceType.PUBLIC_SPACE); // Используем соответствующий тип
            floor.addSpace(defaultSpace);
        }
    }

    private void selectDefaultSpace(Floor floor) {
        for (int i = 0; i < spaceListModel.size(); i++) {
            if ("-".equals(spaceListModel.get(i).getIdentifier())) {
                spaceList.setSelectedIndex(i);
                break;
            }
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
        floorListModel.clear();
        building.getFloors().forEach(floorListModel::addElement);

        if (!floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
            updateSpaceList();
            updateRoomList();
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

    private void saveVentilationCalculations() {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            Arrays.stream(((MainFrame) mainFrame).getTabbedPane().getComponents())
                    .filter(tab -> tab instanceof VentilationTab)
                    .findFirst()
                    .ifPresent(tab -> ((VentilationTab) tab).saveCalculationsToModel());
        }
    }

    private String showInputDialog(String title, String message, String initialValue) {
        JPanel panel = new JPanel(new GridLayout(1, 2, 5, 5));
        panel.add(new JLabel(message));
        JTextField field = new JTextField(initialValue);
        panel.add(field);

        int result = JOptionPane.showConfirmDialog(
                this, panel, title, JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);

        return (result == JOptionPane.OK_OPTION) ? field.getText() : null;
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
        updateRadiationTab(building);
        if (!floorListModel.isEmpty()) floorList.setSelectedIndex(0);

        // Инициализация слушателей после создания компонентов
        floorList.addListSelectionListener(evt -> {
            if (!evt.getValueIsAdjusting()) updateSpaceList();
        });
        spaceList.addListSelectionListener(evt -> {
            if (!evt.getValueIsAdjusting()) updateRoomList();
        });
    }
}