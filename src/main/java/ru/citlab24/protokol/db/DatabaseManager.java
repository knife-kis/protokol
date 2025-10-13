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
                    "position INT)");

            // миграции (безопасны, если столбцы уже есть)
            addColumnIfMissing(stmt, "floor", "section_index", "INT");
            addColumnIfMissing(stmt, "floor", "position", "INT");
            addColumnIfMissing(stmt, "space", "position", "INT");
            addColumnIfMissing(stmt, "room",  "position", "INT");
            addColumnIfMissing(stmt, "room",  "volume", "DOUBLE");
            addColumnIfMissing(stmt, "room",  "ventilation_channels", "INT");
            addColumnIfMissing(stmt, "room",  "ventilation_section_area", "DOUBLE");
            addColumnIfMissing(stmt, "room",  "is_selected", "BOOLEAN DEFAULT FALSE");
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


    private static void deleteBuildingData(int buildingId) throws SQLException {
        // Сначала удаляем комнаты
        try (PreparedStatement stmt = connection.prepareStatement(
                "DELETE FROM room WHERE space_id IN (SELECT id FROM space WHERE floor_id IN (SELECT id FROM floor WHERE building_id = ?))"
        )) {
            stmt.setInt(1, buildingId);
            stmt.executeUpdate();
        }

        // Затем удаляем помещения
        try (PreparedStatement stmt = connection.prepareStatement(
                "DELETE FROM space WHERE floor_id IN (SELECT id FROM floor WHERE building_id = ?)"
        )) {
            stmt.setInt(1, buildingId);
            stmt.executeUpdate();
        }

        // Затем удаляем этажи
        try (PreparedStatement stmt = connection.prepareStatement(
                "DELETE FROM floor WHERE building_id = ?"
        )) {
            stmt.setInt(1, buildingId);
            stmt.executeUpdate();
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
        String sql = "INSERT INTO space (floor_id, identifier, type, position) VALUES (?, ?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, floorId);
            stmt.setString(2, space.getIdentifier());
            stmt.setString(3, space.getType().name());
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


    private static void saveRoom(int spaceId, Room room) throws SQLException {
        logger.debug("Сохранение комнаты: {}", room.getName());
        String sql = "INSERT INTO room (space_id, name, volume, ventilation_channels, ventilation_section_area, is_selected, position) " +
                "VALUES (?, ?, ?, ?, ?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, spaceId);
            stmt.setString(2, room.getName());
            if (room.getVolume() == null) stmt.setNull(3, Types.DOUBLE); else stmt.setDouble(3, room.getVolume());
            stmt.setInt(4, room.getVentilationChannels());
            stmt.setDouble(5, room.getVentilationSectionArea());
            stmt.setBoolean(6, room.isSelected());
            stmt.setInt(7, room.getPosition());
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
                if (floor.getName() == null) {
                    floor.setName(floor.getType().title + " " + floor.getNumber());
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
                    room.setSelected(rs.getBoolean("is_selected"));
                    double volume = rs.getDouble("volume");
                    if (!rs.wasNull()) room.setVolume(volume);
                    room.setVentilationChannels(rs.getInt("ventilation_channels"));
                    room.setVentilationSectionArea(rs.getDouble("ventilation_section_area"));
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

                    // Обработка NULL значения для объема
                    double volume = rs.getDouble("volume");
                    room.setVolume(rs.wasNull() ? null : volume);

                    room.setVentilationChannels(rs.getInt("ventilation_channels"));
                    room.setVentilationSectionArea(rs.getDouble("ventilation_section_area"));
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
}