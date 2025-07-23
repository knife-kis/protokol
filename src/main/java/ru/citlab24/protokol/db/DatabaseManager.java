package ru.citlab24.protokol.db;

import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class DatabaseManager {
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

            stmt.execute("CREATE TABLE IF NOT EXISTS floor (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "building_id INT," +
                    "number VARCHAR(50)," +
                    "type VARCHAR(50))");

            stmt.execute("CREATE TABLE IF NOT EXISTS space (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "floor_id INT," +
                    "identifier VARCHAR(50)," +
                    "type VARCHAR(50))");

            stmt.execute("CREATE TABLE IF NOT EXISTS room (" +
                    "id INT AUTO_INCREMENT PRIMARY KEY," +
                    "space_id INT," +
                    "name VARCHAR(255)," +
                    "volume DOUBLE," + // Разрешено NULL значение
                    "ventilation_channels INT," +
                    "ventilation_section_area DOUBLE)");

            // Проверка столбцов остается без изменений
            addColumnIfMissing(stmt, "room", "volume", "DOUBLE");
            addColumnIfMissing(stmt, "room", "ventilation_channels", "INT");
            addColumnIfMissing(stmt, "room", "ventilation_section_area", "DOUBLE");
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
                    for (Floor floor : building.getFloors()) {
                        saveFloor(buildingId, floor);
                    }
                }
            }
        }
    }

//    private static void insertBuilding(Building building) throws SQLException {
//        try (PreparedStatement stmt = connection.prepareStatement(
//                "INSERT INTO building (name) VALUES (?)",
//                Statement.RETURN_GENERATED_KEYS
//        )) {
//            stmt.setString(1, building.getName());
//            stmt.executeUpdate();
//
//            try (ResultSet rs = stmt.getGeneratedKeys()) {
//                if (rs.next()) {
//                    int buildingId = rs.getInt(1);
//                    building.setId(buildingId);
//                    for (Floor floor : building.getFloors()) {
//                        saveFloor(buildingId, floor);
//                    }
//                }
//            }
//        }
//    }

//    private static void updateBuilding(Building building) throws SQLException {
//        // Обновляем название здания
//        try (PreparedStatement stmt = connection.prepareStatement(
//                "UPDATE building SET name = ? WHERE id = ?"
//        )) {
//            stmt.setString(1, building.getName());
//            stmt.setInt(2, building.getId());
//            stmt.executeUpdate();
//        }
//
//        // Удаляем старые данные и сохраняем заново
//        deleteBuildingData(building.getId());
//
//        // Сохраняем обновленные данные
//        for (Floor floor : building.getFloors()) {
//            saveFloor(building.getId(), floor);
//        }
//    }
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
        String sql = "INSERT INTO floor (building_id, number, type) VALUES (?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, buildingId);
            stmt.setString(2, floor.getNumber());
            stmt.setString(3, floor.getType().name());
            stmt.executeUpdate();

            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) {
                    int floorId = rs.getInt(1);
                    for (Space space : floor.getSpaces()) {
                        saveSpace(floorId, space);
                    }
                }
            }
        }
    }

    private static void saveSpace(int floorId, Space space) throws SQLException {
        String sql = "INSERT INTO space (floor_id, identifier, type) VALUES (?, ?, ?)";
        try (PreparedStatement stmt = connection.prepareStatement(sql, Statement.RETURN_GENERATED_KEYS)) {
            stmt.setInt(1, floorId);
            stmt.setString(2, space.getIdentifier());
            stmt.setString(3, space.getType().name());
            stmt.executeUpdate();

            try (ResultSet rs = stmt.getGeneratedKeys()) {
                if (rs.next()) {
                    int spaceId = rs.getInt(1);
                    for (Room room : space.getRooms()) {
                        saveRoom(spaceId, room);
                    }
                }
            }
        }
    }

    private static void saveRoom(int spaceId, Room room) throws SQLException {
        String sql = "INSERT INTO room (space_id, name, volume, ventilation_channels, ventilation_section_area) " +
                "VALUES (?, ?, ?, ?, ?)";

        try (PreparedStatement stmt = connection.prepareStatement(sql)) {
            stmt.setInt(1, spaceId);
            stmt.setString(2, room.getName());

            // Обработка NULL значения для объема
            if (room.getVolume() != null) {
                stmt.setDouble(3, room.getVolume());
            } else {
                stmt.setNull(3, Types.DOUBLE);
            }

            stmt.setInt(4, room.getVentilationChannels());
            stmt.setDouble(5, room.getVentilationSectionArea());
            stmt.executeUpdate();
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
                System.out.println("Загружаем здание: " + building.getName());
                loadFloors(building, buildingId);
            }
        }
        return building;
    }

    private static void loadFloors(Building building, int buildingId) throws SQLException {
        String sql = "SELECT * FROM floor WHERE building_id = " + buildingId;
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Floor floor = new Floor();
                floor.setNumber(rs.getString("number"));
                floor.setType(Floor.FloorType.valueOf(rs.getString("type")));
                System.out.println("Загружен этаж: " + floor.getNumber());
                building.addFloor(floor);
                loadSpaces(floor, rs.getInt("id"));
            }
        }
    }

    private static void loadSpaces(Floor floor, int floorId) throws SQLException {
        String sql = "SELECT * FROM space WHERE floor_id = " + floorId;
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Space space = new Space();
                space.setIdentifier(rs.getString("identifier"));
                space.setType(Space.SpaceType.valueOf(rs.getString("type")));
                floor.addSpace(space);
                loadRooms(space, rs.getInt("id"));
            }
        }
    }

    private static void loadRooms(Space space, int spaceId) throws SQLException {
        String sql = "SELECT * FROM room WHERE space_id = " + spaceId;
        try (Statement stmt = connection.createStatement();
             ResultSet rs = stmt.executeQuery(sql)) {
            while (rs.next()) {
                Room room = new Room();
                room.setName(rs.getString("name"));

                // Обработка NULL значения для объема
                double volumeValue = rs.getDouble("volume");
                if (rs.wasNull()) {
                    room.setVolume(null);
                } else {
                    room.setVolume(volumeValue);
                }

                room.setVentilationChannels(rs.getInt("ventilation_channels"));
                room.setVentilationSectionArea(rs.getDouble("ventilation_section_area"));
                space.addRoom(room);
            }
        }
    }
    public static List<Building> getBuildings() {
        try {
            return getAllBuildings();
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Ошибка загрузки зданий: " + e.getMessage(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
            return new ArrayList<>();
        }
    }

    public static List<Floor> getFloors(int buildingId) {
        List<Floor> floors = new ArrayList<>();
        String sql = "SELECT * FROM floor WHERE building_id = ?";
        try (PreparedStatement stmt = connection.prepareStatement(sql)) {
            stmt.setInt(1, buildingId);
            try (ResultSet rs = stmt.executeQuery()) {
                while (rs.next()) {
                    Floor floor = new Floor();
                    floor.setId(rs.getInt("id"));
                    floor.setNumber(rs.getString("number"));
                    floor.setType(Floor.FloorType.valueOf(rs.getString("type")));
                    floors.add(floor);
                }
            }
        } catch (SQLException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null,
                    "Ошибка загрузки этажей: " + e.getMessage(),
                    "Database Error", JOptionPane.ERROR_MESSAGE);
        }
        return floors;
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