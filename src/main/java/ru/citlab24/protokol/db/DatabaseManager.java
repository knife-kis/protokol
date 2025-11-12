package ru.citlab24.protokol.db;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.sql.*;
import java.util.*;

public class DatabaseManager {
    private static final Logger logger = LoggerFactory.getLogger(DatabaseManager.class);
    private static final String CONFIG_FILE = "/db.properties";
    private static Connection connection;

    static {
        try {
            Properties prop = new Properties();
            prop.load(DatabaseManager.class.getResourceAsStream(CONFIG_FILE));

            // Создаем встроенное соединение с H2
            connection = DriverManager.getConnection(
                    prop.getProperty("db.url"),
                    prop.getProperty("db.user"),
                    prop.getProperty("db.password")
            );

            createTables(); // Создаем таблицы при старте
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Ошибка инициализации базы данных: " + e.getMessage(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static void createTables() throws SQLException {
        try (Statement stmt = connection.createStatement()) {
            stmt.execute("CREATE TABLE IF NOT EXISTS building (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "name VARCHAR(255))");

            stmt.execute("CREATE TABLE IF NOT EXISTS section (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "building_id INT," +
                    "name VARCHAR(255)," +
                    "position INT)");

            stmt.execute("CREATE TABLE IF NOT EXISTS floor (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "building_id INT," +
                    "number VARCHAR(50)," +
                    "type VARCHAR(50)," +
                    "section_index INT," +
                    "position INT)");

            stmt.execute("CREATE TABLE IF NOT EXISTS space (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "floor_id INT," +
                    "identifier VARCHAR(255)," +
                    "type VARCHAR(50)," +
                    "position INT)");

            stmt.execute("CREATE TABLE IF NOT EXISTS room (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "space_id INT," +
                    "name VARCHAR(255)," +
                    "volume DOUBLE," +
                    "ventilation_channels INT," +
                    "ventilation_section_area DOUBLE," +
                    "is_selected BOOLEAN DEFAULT FALSE," +
                    "artificial_selected BOOLEAN DEFAULT FALSE," +
                    "external_walls_count INT," +
                    "microclimate_selected BOOLEAN DEFAULT FALSE," +
                    "radiation_selected  BOOLEAN DEFAULT FALSE," +
                    "street_left_max   DOUBLE," +
                    "street_center_min DOUBLE," +
                    "street_right_max  DOUBLE," +
                    "street_bottom_min DOUBLE," +
                    "position INT)");

            // НОВОЕ: отдельная таблица «Шумы» (без изменений моделей)
            stmt.execute("CREATE TABLE IF NOT EXISTS noise_settings (" +
                    "room_id INT PRIMARY KEY," +
                    "measure BOOLEAN," +
                    "lift BOOLEAN," +
                    "vent BOOLEAN," +
                    "heat_curtain BOOLEAN," +
                    "itp BOOLEAN," +
                    "pns BOOLEAN," +
                    "electrical BOOLEAN," +
                    "auto_src BOOLEAN," +
                    "zum BOOLEAN)");

            // миграции

            addColumnIfMissing(stmt, "floor", "section_index", "INT");
            addColumnIfMissing(stmt, "floor", "position", "INT");
            addColumnIfMissing(stmt, "space", "position", "INT");
            addColumnIfMissing(stmt, "room",  "position", "INT");
            addColumnIfMissing(stmt, "room",  "volume", "DOUBLE");
            addColumnIfMissing(stmt, "room",  "ventilation_channels", "INT");
            addColumnIfMissing(stmt, "room",  "ventilation_section_area", "DOUBLE");
            addColumnIfMissing(stmt, "room",  "is_selected", "BOOLEAN DEFAULT FALSE");
            addColumnIfMissing(stmt, "room",  "external_walls_count", "INT");
            addColumnIfMissing(stmt, "room",  "microclimate_selected",  "BOOLEAN DEFAULT FALSE");
            addColumnIfMissing(stmt, "room",  "radiation_selected",     "BOOLEAN DEFAULT FALSE");
            addColumnIfMissing(stmt, "room",  "artificial_selected",    "BOOLEAN DEFAULT FALSE");
            addColumnIfMissing(stmt, "room", "street_left_max",   "DOUBLE");
            addColumnIfMissing(stmt, "room", "street_center_min", "DOUBLE");
            addColumnIfMissing(stmt, "room", "street_right_max",  "DOUBLE");
            addColumnIfMissing(stmt, "room", "street_bottom_min", "DOUBLE");
            addColumnIfMissing(stmt, "room",  "ventilation_duct_shape", "VARCHAR(16)");
            addColumnIfMissing(stmt, "room",  "ventilation_width",      "DOUBLE");
        }
    }

    private static void addColumnIfMissing(Statement stmt, String table, String column, String type)
            throws SQLException {
        ResultSet rs = stmt.executeQuery(
                "SELECT COUNT(*) FROM INFORMATION_SCHEMA.COLUMNS " +
                        "WHERE TABLE_NAME = '" + table.toUpperCase() + "' " +
                        "AND COLUMN_NAME = '" + column.toUpperCase() + "'"
        );
        rs.next();
        if (rs.getInt(1) == 0) {
            stmt.execute("ALTER TABLE " + table + " ADD COLUMN " + column + " " + type);
        }
    }

    public static void saveBuilding(Building building) throws SQLException {
        try (PreparedStatement stmt = connection.prepareStatement(
                "INSERT INTO building (name) VALUES (?)",
                Statement.RETURN_GENERATED_KEYS
        )) {
            stmt.setString(1, building.getName());
            stmt.executeUpdate();

            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) {
                    int buildingId = rs.getInt(1);
                    building.setId(buildingId);

                    // 1) Сначала секции
                    saveSections(buildingId, building.getSections());

                    // 2) Затем этажи (у них уже корректный section_index)
                    for (Floor floor : building.getFloors()) {
                        saveFloor(buildingId, floor);
                    }
                }
            }
        }
    }
    public static void deleteBuilding(int buildingId) throws SQLException {
        // Удаляем комнаты данного здания
        try (PreparedStatement ps = connection.prepareStatement(
                "DELETE FROM room WHERE space_id IN (" +
                        "  SELECT s.id FROM space s JOIN floor f ON s.floor_id = f.id WHERE f.building_id = ?" +
                        ")")) {
            ps.setInt(1, buildingId);
            ps.executeUpdate();
        }

        // Удаляем помещения этого здания
        try (PreparedStatement ps = connection.prepareStatement(
                "DELETE FROM space WHERE floor_id IN (" +
                        "  SELECT id FROM floor WHERE building_id = ?" +
                        ")")) {
            ps.setInt(1, buildingId);
            ps.executeUpdate();
        }

        // Удаляем этажи этого здания
        try (PreparedStatement ps = connection.prepareStatement(
                "DELETE FROM floor WHERE building_id = ?")) {
            ps.setInt(1, buildingId);
            ps.executeUpdate();
        }

        // Удаляем секции этого здания (если таблица section есть)
        try (PreparedStatement ps = connection.prepareStatement(
                "DELETE FROM section WHERE building_id = ?")) {
            ps.setInt(1, buildingId);
            ps.executeUpdate();
        } catch (SQLException ignore) {
            // на случай если таблицы секций нет в старой БД
        }

        // Удаляем сам объект здания
        try (PreparedStatement ps = connection.prepareStatement(
                "DELETE FROM building WHERE id = ?")) {
            ps.setInt(1, buildingId);
            ps.executeUpdate();
        }
    }


    public static List<Building> getAllBuildings() throws SQLException {
        List<Building> buildings = new ArrayList<>();
        String sql = "SELECT * FROM building";
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Building building = new Building();
                building.setId(rs.getInt("id"));
                building.setName(rs.getString("name"));
                buildings.add(building);
            }
        }
        return buildings;
    }

    private static void saveFloor(int buildingId, Floor floor) throws SQLException {
        String sql = "INSERT INTO floor (building_id, number, type, section_index, position) VALUES (?,?,?,?,?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, buildingId);
            stmt.setString(2, floor.getNumber());
            stmt.setString(3, floor.getType().name());
            stmt.setInt(4, floor.getSectionIndex());
            stmt.setInt(5, floor.getPosition());
            stmt.executeUpdate();

            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) {
                    int floorId = rs.getInt(1);
                    floor.setId(floorId);
                    for (Space space : floor.getSpaces()) {
                        saveSpace(floorId, space);
                    }
                }
            }
        }
    }

    private static void saveSections(int buildingId, List<Section> sections) throws SQLException {
        if (sections == null) return;
        String sql = "INSERT INTO section (building_id, name, position) VALUES (?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            for (Section s : sections) {
                stmt.setInt(1, buildingId);
                stmt.setString(2, s.getName());
                stmt.setInt(3, s.getPosition());
                stmt.executeUpdate();
                try (ResultSet rs = stmt.getGeneratedKeys()) {
                    if (rs.next()) s.setId(rs.getInt(1));
                }
            }
        }
    }

    private static void saveSpace(int floorId, Space space) throws SQLException {
        // Если этаж уличный — принудительно сохраняем помещения как OUTDOOR
        Floor.FloorType parentType = getFloorType(floorId);
        Space.SpaceType effectiveType = space.getType();
        if (parentType == Floor.FloorType.STREET && effectiveType != Space.SpaceType.OUTDOOR) {
            logger.warn("Этаж {} — STREET. Тип помещения '{}' переопределён на OUTDOOR.",
                    floorId, effectiveType);
            effectiveType = Space.SpaceType.OUTDOOR;
        }

        String sql = "INSERT INTO space (floor_id, identifier, type, position) VALUES (?, ?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, floorId);
            stmt.setString(2, space.getIdentifier());
            stmt.setString(3, effectiveType.name());
            stmt.setInt(4, space.getPosition());
            stmt.executeUpdate();

            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) {
                    int spaceId = rs.getInt(1);
                    space.setId(spaceId);
                    for (Room room : space.getRooms()) {
                        saveRoom(spaceId, room);
                    }
                }
            }
        }
    }
    private static Floor.FloorType getFloorType(int floorId) throws SQLException {
        String sql = "SELECT type FROM floor WHERE id = ?";
        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            ps.setInt(1, floorId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    String typeStr = rs.getString(1);
                    try {
                        return Floor.FloorType.valueOf(typeStr);
                    } catch (IllegalArgumentException ignore) {
                        logger.warn("Неизвестный тип этажа в БД: '{}', используем PUBLIC по умолчанию", typeStr);
                    }
                }
            }
        }
        return Floor.FloorType.PUBLIC;
    }

    // Стало
    private static void saveRoom(int spaceId, Room room) throws SQLException {
        logger.debug("Сохранение комнаты: {}", room.getName());
        String sql = "INSERT INTO room (" +
                "space_id, name, volume, " +
                "ventilation_channels, ventilation_section_area, ventilation_duct_shape, ventilation_width, " +
                "is_selected, artificial_selected, external_walls_count, " +
                "microclimate_selected, radiation_selected, position" +
                ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";

        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, spaceId);
            stmt.setString(2, room.getName());

            // volume (nullable)
            if (room.getVolume() == null) stmt.setNull(3, Types.DOUBLE);
            else stmt.setDouble(3, room.getVolume());

            // вентиляция (каналы, сечение одного канала — как было)
            stmt.setInt(4, room.getVentilationChannels());
            stmt.setDouble(5, room.getVentilationSectionArea());

            // НОВОЕ: форма и ширина (через рефлексию — совместимо со старыми Room)
            String ductShape = null;
            Double ductWidth = null;
            try {
                Object v = room.getClass().getMethod("getVentilationDuctShape").invoke(room);
                if (v != null) ductShape = String.valueOf(v);
            } catch (Throwable ignore) {}
            try {
                Object v = room.getClass().getMethod("getVentilationWidth").invoke(room);
                if (v instanceof Number n) ductWidth = n.doubleValue();
                else if (v != null) ductWidth = Double.valueOf(v.toString());
            } catch (Throwable ignore) {}

            if (ductShape == null || ductShape.isBlank()) stmt.setNull(6, Types.VARCHAR);
            else stmt.setString(6, ductShape);

            if (ductWidth == null) stmt.setNull(7, Types.DOUBLE);
            else stmt.setDouble(7, ductWidth);

            // КЕО
            stmt.setBoolean(8, room.isSelected());

            // Искусственное освещение (рефлексия как было)
            boolean artificial = false;
            try {
                Object v = room.getClass().getMethod("isArtificialSelected").invoke(room);
                if (v instanceof Boolean) artificial = (Boolean) v;
            } catch (Throwable ignore) {
                try {
                    Object v2 = room.getClass().getMethod("isArtificialLightingSelected").invoke(room);
                    if (v2 instanceof Boolean) artificial = (Boolean) v2;
                } catch (Throwable ignore2) {
                    try {
                        Object v3 = room.getClass().getMethod("getArtificialLightingSelected").invoke(room);
                        if (v3 instanceof Boolean) artificial = (Boolean) v3;
                    } catch (Throwable ignore3) {}
                }
            }
            stmt.setBoolean(9, artificial);

            // наружные стены (nullable)
            Integer walls = room.getExternalWallsCount();
            if (walls == null) stmt.setNull(10, Types.INTEGER); else stmt.setInt(10, walls);

            // микроклимат / радиация
            stmt.setBoolean(11, room.isMicroclimateSelected());
            stmt.setBoolean(12, room.isRadiationSelected());

            // порядок
            stmt.setInt(13, room.getPosition());

            stmt.executeUpdate();
            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) room.setId(rs.getInt(1));
            }
        }
    }

    public static Building loadBuilding(int buildingId) throws SQLException {
        Building building = new Building();
        String sql = "SELECT * FROM building WHERE id = " + buildingId;
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            if (rs.next()) {
                building.setId(rs.getInt("id"));
                building.setName(rs.getString("name"));
                loadSections(building, buildingId);
                loadFloors(building, buildingId);
            }
        }
        return building;
    }

    private static void loadFloors(Building building, int buildingId) throws SQLException {
        String sql = "SELECT * FROM floor WHERE building_id = " + buildingId +
                " ORDER BY section_index, COALESCE(position,0), id";
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Floor floor = new Floor();
                floor.setId(rs.getInt("id"));
                floor.setNumber(rs.getString("number"));
                floor.setType(Floor.FloorType.valueOf(rs.getString("type")));
                floor.setSectionIndex(Math.max(0, rs.getInt("section_index")));
                floor.setPosition(rs.getInt("position"));

                // Автоимя: для STREET — «Улица» (без номера, если он пуст),
                // для остальных — "<название типа> <номер>" (номер добавляем только если он не пуст)
                String num = floor.getNumber();
                boolean hasNum = (num != null && !num.isBlank());
                String autoName;
                if (floor.getType() == Floor.FloorType.STREET) {
                    autoName = hasNum ? ("Улица " + num) : "Улица";
                } else {
                    autoName = floor.getType().title + (hasNum ? (" " + num) : "");
                }
                if (floor.getName() == null || floor.getName().isBlank()) {
                    floor.setName(autoName);
                }

                building.addFloor(floor);
                loadSpaces(floor, floor.getId());
            }
        }
    }

    private static void loadSpaces(Floor floor, int floorId) throws SQLException {
        String sql = "SELECT * FROM space WHERE floor_id = " + floorId + " ORDER BY COALESCE(position,0), id";
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Space space = new Space();
                space.setId(rs.getInt("id"));
                space.setIdentifier(rs.getString("identifier"));
                space.setType(Space.SpaceType.valueOf(rs.getString("type")));
                space.setPosition(rs.getInt("position"));
                floor.addSpace(space);
                loadRooms(space, space.getId());
            }
        }
    }

    // Стало
    private static void loadRooms(Space space, int spaceId) throws SQLException {
        logger.debug("Загрузка комнат для помещения ID: {}", spaceId);
        String sql = "SELECT * FROM room WHERE space_id = ? ORDER BY COALESCE(position,0), id";
        try (PreparedStatement stmt = connection.prepareStatement(sql)) {
            stmt.setInt(1, spaceId);
            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    Room room = new Room();
                    room.setId(rs.getInt("id"));
                    room.setName(rs.getString("name"));

                    // Объём (nullable)
                    double volume = rs.getDouble("volume");
                    room.setVolume(rs.wasNull() ? null : volume);

                    // Вентиляция — каналы и сечение (как было)
                    room.setVentilationChannels(rs.getInt("ventilation_channels"));
                    room.setVentilationSectionArea(rs.getDouble("ventilation_section_area"));

                    // НОВОЕ: форма и ширина (через сеттеры, если есть; иначе тихо пропускаем)
                    try {
                        String shape = null;
                        try { shape = rs.getString("ventilation_duct_shape"); } catch (SQLException ignore) {}
                        if (shape != null) {
                            try { room.getClass().getMethod("setVentilationDuctShape", String.class).invoke(room, shape); }
                            catch (Throwable ignore) {}
                        }
                        Double width = null;
                        try {
                            Object w = rs.getObject("ventilation_width");
                            if (w instanceof Number n) width = n.doubleValue();
                        } catch (SQLException ignore) {}
                        if (width != null) {
                            // сначала Double-версией, затем double-версией
                            try { room.getClass().getMethod("setVentilationWidth", Double.class).invoke(room, width); }
                            catch (NoSuchMethodException ex) {
                                try { room.getClass().getMethod("setVentilationWidth", double.class).invoke(room, width.doubleValue()); }
                                catch (Throwable ignore) {}
                            } catch (Throwable ignore) {}
                        }
                    } catch (Throwable ignore) {}

                    // КЕО
                    try { room.setSelected(rs.getBoolean("is_selected")); } catch (SQLException ignore) {}

                    // Искусственное
                    try {
                        boolean art = rs.getBoolean("artificial_selected");
                        try { room.getClass().getMethod("setArtificialSelected", boolean.class).invoke(room, art); }
                        catch (ReflectiveOperationException ignore) {}
                    } catch (SQLException ignore) {}

                    // Наружные стены
                    Object wallsObj = rs.getObject("external_walls_count");
                    room.setExternalWallsCount(wallsObj == null ? null : ((Number) wallsObj).intValue());

                    // МК / Радиация
                    try { room.setMicroclimateSelected(rs.getBoolean("microclimate_selected")); } catch (SQLException ignore) {}
                    try { room.setRadiationSelected(rs.getBoolean("radiation_selected")); } catch (SQLException ignore) {}

                    // Порядок
                    room.setPosition(rs.getInt("position"));

                    space.addRoom(room);
                }
            }
        }
    }

    private static void loadSections(Building building, int buildingId) throws SQLException {
        String sql = "SELECT * FROM section WHERE building_id = " + buildingId + " ORDER BY position";
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            List<Section> sections = new ArrayList<>();
            while (rs.next()) {
                Section s = new Section();
                s.setId(rs.getInt("id"));
                s.setName(rs.getString("name"));
                s.setPosition(rs.getInt("position"));
                sections.add(s);
            }
            if (sections.isEmpty()) sections.add(new Section("Секция 1", 0));
            building.setSections(sections);
        }
    }

    // Стало
    public static List<Room> getRooms(int floorId) {
        List<Room> rooms = new ArrayList<>();
        String sql = "SELECT r.* FROM room r " +
                "JOIN space s ON r.space_id = s.id " +
                "JOIN floor f ON s.floor_id = f.id " +
                "WHERE f.id = ?";
        try (PreparedStatement stmt = connection.prepareStatement(sql)) {
            stmt.setInt(1, floorId);
            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    Room room = new Room();
                    room.setId(rs.getInt("id"));
                    room.setName(rs.getString("name"));

                    // Объём
                    double volume = rs.getDouble("volume");
                    room.setVolume(rs.wasNull() ? null : volume);

                    // Вентиляция
                    room.setVentilationChannels(rs.getInt("ventilation_channels"));
                    room.setVentilationSectionArea(rs.getDouble("ventilation_section_area"));

                    // НОВОЕ: форма/ширина
                    try {
                        String shape = null;
                        try { shape = rs.getString("ventilation_duct_shape"); } catch (SQLException ignore) {}
                        if (shape != null) {
                            try { room.getClass().getMethod("setVentilationDuctShape", String.class).invoke(room, shape); }
                            catch (Throwable ignore) {}
                        }
                        Double width = null;
                        try {
                            Object w = rs.getObject("ventilation_width");
                            if (w instanceof Number n) width = n.doubleValue();
                        } catch (SQLException ignore) {}
                        if (width != null) {
                            try { room.getClass().getMethod("setVentilationWidth", Double.class).invoke(room, width); }
                            catch (NoSuchMethodException ex) {
                                try { room.getClass().getMethod("setVentilationWidth", double.class).invoke(room, width.doubleValue()); }
                                catch (Throwable ignore) {}
                            } catch (Throwable ignore) {}
                        }
                    } catch (Throwable ignore) {}

                    // КЕО
                    try { room.setSelected(rs.getBoolean("is_selected")); } catch (SQLException ignore) {}

                    // Искусственное
                    try {
                        boolean art = rs.getBoolean("artificial_selected");
                        try { room.getClass().getMethod("setArtificialSelected", boolean.class).invoke(room, art); }
                        catch (ReflectiveOperationException ignore) {}
                    } catch (SQLException ignore) {}

                    // Наружные стены
                    try {
                        Object wallsObj = rs.getObject("external_walls_count");
                        room.setExternalWallsCount(wallsObj == null ? null : ((Number) wallsObj).intValue());
                    } catch (SQLException ignore) {}

                    // МК / Радиация
                    try { room.setMicroclimateSelected(rs.getBoolean("microclimate_selected")); } catch (SQLException ignore) {}
                    try { room.setRadiationSelected(rs.getBoolean("radiation_selected")); } catch (SQLException ignore) {}

                    rooms.add(room);
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Ошибка загрузки помещений: " + e.getMessage(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
        }
        return rooms;
    }

    public static void updateArtificialSelections(Building b, Map<String, Boolean> byKey) throws SQLException {
        if (b == null || byKey == null || byKey.isEmpty()) return;
        String sql = "UPDATE room SET artificial_selected=? WHERE id=?";

        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            for (Floor f : b.getFloors()) {
                for (Space s : f.getSpaces()) {
                    for (Room r : s.getRooms()) {
                        String key = makeKey(f, s, r);
                        Boolean v = byKey.get(key);
                        if (v == null) continue;
                        ps.setBoolean(1, v);
                        ps.setInt(2, r.getId());
                        ps.addBatch();
                    }
                }
            }
            ps.executeBatch();
        }
    }

    public static Map<String, Boolean> loadArtificialSelectionsByKey(int buildingId) throws SQLException {
        Map<String, Boolean> res = new HashMap<>();
        String sql =
                "SELECT f.section_index, f.number, s.identifier, r.name, r.artificial_selected " +
                        "FROM room r " +
                        "JOIN space s ON r.space_id = s.id " +
                        "JOIN floor f ON s.floor_id = f.id " +
                        "WHERE f.building_id = ?";

        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            ps.setInt(1, buildingId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String key = rs.getInt(1) + "|" +
                            ns(rs.getString(2)) + "|" +
                            ns(rs.getString(3)) + "|" +
                            ns(rs.getString(4));
                    res.put(key, rs.getBoolean(5));
                }
            }
        }
        return res;
    }

    private static String makeKey(Floor f, Space s, Room r) {
        return f.getSectionIndex() + "|" + ns(f.getNumber()) + "|" + ns(s.getIdentifier()) + "|" + ns(r.getName());
    }
    private static String ns(String s) { return (s == null) ? "" : s.trim(); }
    public static void updateStreetLightingValues(Building b, java.util.Map<String, Double[]> byKey) throws SQLException {
        if (b == null || byKey == null || byKey.isEmpty()) return;
        String sql = "UPDATE room SET street_left_max=?, street_center_min=?, street_right_max=?, street_bottom_min=? WHERE id=?";
        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            for (Floor f : b.getFloors()) {
                for (Space s : f.getSpaces()) {
                    for (Room r : s.getRooms()) {
                        String key = makeKey(f, s, r);
                        Double[] v = byKey.get(key);
                        if (v == null) continue;

                        if (v.length > 0 && v[0] != null) ps.setDouble(1, v[0]); else ps.setNull(1, Types.DOUBLE);
                        if (v.length > 1 && v[1] != null) ps.setDouble(2, v[1]); else ps.setNull(2, Types.DOUBLE);
                        if (v.length > 2 && v[2] != null) ps.setDouble(3, v[2]); else ps.setNull(3, Types.DOUBLE);
                        if (v.length > 3 && v[3] != null) ps.setDouble(4, v[3]); else ps.setNull(4, Types.DOUBLE);

                        ps.setInt(5, r.getId());
                        ps.addBatch();
                    }
                }
            }
            ps.executeBatch();
        }
    }
    public static java.util.Map<String, Double[]> loadStreetLightingValuesByKey(int buildingId) throws SQLException {
        java.util.Map<String, Double[]> res = new java.util.HashMap<>();
        String sql =
                "SELECT f.section_index, f.number, s.identifier, r.name, " +
                        "       r.street_left_max, r.street_center_min, r.street_right_max, r.street_bottom_min " +
                        "FROM room r " +
                        "JOIN space s ON r.space_id = s.id " +
                        "JOIN floor f ON s.floor_id = f.id " +
                        "WHERE f.building_id = ?";
        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            ps.setInt(1, buildingId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String key = rs.getInt(1) + "|" + ns(rs.getString(2)) + "|" + ns(rs.getString(3)) + "|" + ns(rs.getString(4));
                    Double v1 = rs.getObject(5) == null ? null : ((Number) rs.getObject(5)).doubleValue();
                    Double v2 = rs.getObject(6) == null ? null : ((Number) rs.getObject(6)).doubleValue();
                    Double v3 = rs.getObject(7) == null ? null : ((Number) rs.getObject(7)).doubleValue();
                    Double v4 = rs.getObject(8) == null ? null : ((Number) rs.getObject(8)).doubleValue();
                    res.put(key, new Double[]{v1, v2, v3, v4});
                }
            }
        }
        return res;
    }

    // ==== «Шумы»: DTO для обмена с вкладкой ====
    public static class NoiseValue {
        public boolean measure;
        public boolean lift;
        public boolean vent;
        public boolean heatCurtain;
        public boolean itp;
        public boolean pns;
        public boolean electrical;
        public boolean autoSrc;
        public boolean zum;
    }

    // Обновить noise_settings по ключам (section|floor|space|room) → room.id нового проекта
    public static void updateNoiseSelections(Building b, Map<String, NoiseValue> byKey) throws SQLException {
        if (b == null || byKey == null || byKey.isEmpty()) return;

        String mergeSql = "MERGE INTO noise_settings (room_id, measure, lift, vent, heat_curtain, itp, pns, electrical, auto_src, zum) " +
                "KEY(room_id) VALUES (?,?,?,?,?,?,?,?,?,?)";

        try (PreparedStatement ps = connection.prepareStatement(mergeSql)) {
            for (Floor f : b.getFloors()) {
                for (Space s : f.getSpaces()) {
                    for (Room r : s.getRooms()) {
                        String key = f.getSectionIndex() + "|" + ns(f.getNumber()) + "|" + ns(s.getIdentifier()) + "|" + ns(r.getName());
                        NoiseValue v = byKey.get(key);
                        if (v == null) continue;

                        ps.setInt(1, r.getId());
                        ps.setBoolean(2, v.measure);
                        ps.setBoolean(3, v.lift);
                        ps.setBoolean(4, v.vent);
                        ps.setBoolean(5, v.heatCurtain);
                        ps.setBoolean(6, v.itp);
                        ps.setBoolean(7, v.pns);
                        ps.setBoolean(8, v.electrical);
                        ps.setBoolean(9, v.autoSrc);
                        ps.setBoolean(10, v.zum);
                        ps.addBatch();
                    }
                }
            }
            ps.executeBatch();
        }
    }

    // Прочитать noise_settings в карту по ключу (section|floor|space|room) для buildingId
    public static Map<String, NoiseValue> loadNoiseSelectionsByKey(int buildingId) throws SQLException {
        Map<String, NoiseValue> res = new HashMap<>();
        String sql = "SELECT f.section_index, f.number, s.identifier, r.name, " +
                " n.measure, n.lift, n.vent, n.heat_curtain, n.itp, n.pns, n.electrical, n.auto_src, n.zum " +
                "FROM room r " +
                "JOIN space s ON r.space_id = s.id " +
                "JOIN floor f ON s.floor_id = f.id " +
                "LEFT JOIN noise_settings n ON n.room_id = r.id " +
                "WHERE f.building_id = ?";

        try (PreparedStatement ps = connection.prepareStatement(sql)) {
            ps.setInt(1, buildingId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    String key = rs.getInt(1) + "|" + ns(rs.getString(2)) + "|" + ns(rs.getString(3)) + "|" + ns(rs.getString(4));
                    NoiseValue nv = new NoiseValue();
                    nv.measure     = rs.getBoolean(5);
                    nv.lift        = rs.getBoolean(6);
                    nv.vent        = rs.getBoolean(7);
                    nv.heatCurtain = rs.getBoolean(8);
                    nv.itp         = rs.getBoolean(9);
                    nv.pns         = rs.getBoolean(10);
                    nv.electrical  = rs.getBoolean(11);
                    nv.autoSrc     = rs.getBoolean(12);
                    nv.zum         = rs.getBoolean(13);
                    res.put(key, nv);
                }
            }
        }
        return res;
    }
    /** Обновить/вставить настройки шума по ключу "sectionIndex|floorNumber|spaceIdentifier|roomName" в рамках здания. */
    public static void updateNoiseValueByKey(Building building, String key, NoiseValue v) {
        if (key == null || v == null) return;
        int buildingId = (building != null) ? building.getId() : 0;

        // Парсим ключ безопасно
        String[] parts = (key == null) ? new String[0] : key.split("\\|", -1);
        String sIdxStr   = (parts.length > 0) ? parts[0].trim() : "0";
        String floorNum  = (parts.length > 1) ? ns(parts[1])    : "";
        String spaceId   = (parts.length > 2) ? ns(parts[2])    : "";
        String roomName  = (parts.length > 3) ? ns(parts[3])    : "";
        int sectionIndex = 0;
        try { sectionIndex = Integer.parseInt(sIdxStr); } catch (NumberFormatException ignore) { sectionIndex = 0; }

        try {
            // Находим room.id по ключу (с фильтром по building_id, если он известен)
            String sql = "SELECT r.id " +
                    "FROM room r " +
                    "JOIN space s ON r.space_id = s.id " +
                    "JOIN floor f ON s.floor_id = f.id " +
                    (buildingId > 0 ? "WHERE f.building_id = ? AND " : "WHERE ") +
                    "f.section_index = ? AND COALESCE(f.number,'') = ? AND " +
                    "COALESCE(s.identifier,'') = ? AND COALESCE(r.name,'') = ? " +
                    "LIMIT 1";

            try (PreparedStatement ps = connection.prepareStatement(sql)) {
                int idx = 1;
                if (buildingId > 0) ps.setInt(idx++, buildingId);
                ps.setInt(idx++, sectionIndex);
                ps.setString(idx++, floorNum);
                ps.setString(idx++, spaceId);
                ps.setString(idx++, roomName);

                try (ResultSet rs = ps.executeQuery()) {
                    if (!rs.next()) {
                        logger.warn("updateNoiseValueByKey: не найден room по ключу '{}'", key);
                        return;
                    }
                    int roomId = rs.getInt(1);

                    String merge = "MERGE INTO noise_settings " +
                            "(room_id, measure, lift, vent, heat_curtain, itp, pns, electrical, auto_src, zum) " +
                            "KEY(room_id) VALUES (?,?,?,?,?,?,?,?,?,?)";

                    try (PreparedStatement pm = connection.prepareStatement(merge)) {
                        pm.setInt(1, roomId);
                        pm.setBoolean(2, v.measure);
                        pm.setBoolean(3, v.lift);
                        pm.setBoolean(4, v.vent);
                        pm.setBoolean(5, v.heatCurtain);
                        pm.setBoolean(6, v.itp);
                        pm.setBoolean(7, v.pns);
                        pm.setBoolean(8, v.electrical);
                        pm.setBoolean(9, v.autoSrc);
                        pm.setBoolean(10, v.zum);
                        pm.executeUpdate();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Ошибка обновления шумов: " + e.getMessage(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
        }
    }
}