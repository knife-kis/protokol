package ru.citlab24.protokol.tabs.buildingTab;

import ru.citlab24.protokol.tabs.models.*;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** Небольшой сервис «модельных» операций: распознавание типов, копирование этажей/помещений/комнат,
 * нормализация имён, автопроставления, снимки флагов и т. п. Без UI и без зависимостей от вкладок. */
public final class BuildingModelOps {

    private Building building;

    public BuildingModelOps(Building building) {
        this.building = (building != null) ? building : new Building();
    }

    public void setBuilding(Building building) {
        this.building = (building != null) ? building : new Building();
    }

    /* ===================== НОРМАЛИЗАЦИЯ НАЗВАНИЙ ===================== */

    private static final java.util.regex.Pattern TAIL_PLO =
            java.util.regex.Pattern.compile(
                    "\\s*[,;]?\\s*площад[ьяиуе][^\\d]*\\d+[\\d.,]*\\s*(?:кв\\.?\\s*м|м\\s*²|м2|м\\^2)\\.?\\s*$",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    private static final java.util.regex.Pattern TAIL_AXES =
            java.util.regex.Pattern.compile(
                    "\\s*[,;]?\\s*в\\s+осях\\s+[\\p{L}\\p{N}]+(?:\\s*[-—–]\\s*[\\p{L}\\p{N}]+)?\\s*$",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    private static final java.util.regex.Pattern TAIL_PAREN =
            java.util.regex.Pattern.compile("\\s*\\((?:[^)(]+|\\([^)(]*\\))*\\)\\s*$");

    public static String normalizeRoomBaseName(String name) {
        if (name == null) return "";
        String s = name.replace('\u00A0',' ').trim();
        s = s.replaceAll("\\s+", " ");

        boolean changed;
        do {
            changed = false;
            java.util.regex.Matcher mp = TAIL_PAREN.matcher(s);
            if (mp.find()) { s = s.substring(0, mp.start()).trim(); changed = true; }
        } while (changed);

        s = TAIL_AXES.matcher(s).replaceAll("");
        s = TAIL_PLO.matcher(s).replaceAll("");
        s = s.replaceAll("[\\s\\.,;:—–-]+$", "").trim();
        return s;
    }

    private static final java.util.regex.Pattern PUB_TAIL_NUM =
            java.util.regex.Pattern.compile("\\s*(?:№\\s*)?\\d+\\s*$");

    public static String normalizePublicBaseName(String name) {
        String s = normalizeRoomBaseName(name);
        s = PUB_TAIL_NUM.matcher(s).replaceAll("");
        s = s.replaceAll("[\\s\\.,;:—–-]+$", "").trim();
        return s;
    }

    private static final java.util.regex.Pattern OFFICE_TAIL_NUM =
            java.util.regex.Pattern.compile("\\s*(?:№\\s*)?\\d+\\s*$");

    public static String normalizeOfficeBaseName(String name) {
        if (name == null) return "";
        String s = name.replace('\u00A0',' ').trim().replaceAll("\\s+", " ");
        s = OFFICE_TAIL_NUM.matcher(s).replaceAll("");
        s = s.replaceAll("[\\s\\.,;:—–-]+$", "").trim();
        return s;
    }

    public static boolean looksLikeSanitary(String name) {
        if (name == null) return false;
        String s = name.toLowerCase(Locale.ROOT);
        return s.contains("сануз") || s.contains("с/у") || s.contains("сан.уз")
                || s.contains("туалет") || s.contains("унитаз")
                || s.contains("ванн") || s.contains("душ")
                || s.contains("уборная") || s.contains("wc") || s.contains("toilet");
    }

    /* ===================== КЛАССИФИКАЦИЯ ТИПОВ ===================== */

    public boolean isPublicFloor(Floor f) {
        if (f == null) return false;
        try {
            String n = (f.getType() != null) ? f.getType().name() : "";
            if ("PUBLIC".equalsIgnoreCase(n) ||
                    "PUBLIC_AREA".equalsIgnoreCase(n) ||
                    "COMMON".equalsIgnoreCase(n) ||
                    "COMMUNITY".equalsIgnoreCase(n)) return true;
        } catch (Throwable ignore) {}
        String title = (f.getType() != null && f.getType().title != null) ? f.getType().title : "";
        return title.toLowerCase(Locale.ROOT).contains("обществен");
    }

    public boolean isOfficeFloor(Floor f) {
        if (f == null) return false;
        try {
            String n = (f.getType() != null) ? f.getType().name() : "";
            if ("OFFICE".equalsIgnoreCase(n)) return true;
        } catch (Throwable ignore) {}
        String title = (f.getType() != null && f.getType().title != null) ? f.getType().title : "";
        String low = title.toLowerCase(Locale.ROOT);
        return low.contains("офис");
    }

    public Floor findFloorForSpace(Space s) {
        if (s == null) return null;
        for (Floor f : building.getFloors()) {
            for (Space each : f.getSpaces()) {
                if (each == s) return f;
            }
        }
        return null;
    }

    public boolean isApartmentSpace(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t != null) {
            String tn = t.name();
            if ("APARTMENT".equalsIgnoreCase(tn) ||
                    "FLAT".equalsIgnoreCase(tn) ||
                    "RESIDENTIAL".equalsIgnoreCase(tn) ||
                    "LIVING".equalsIgnoreCase(tn)) {
                return true;
            }
        }
        String id = s.getIdentifier();
        if (id != null) {
            String low = id.toLowerCase(Locale.ROOT);
            if (low.startsWith("кв") || low.contains("квартира")) return true;
            if (low.startsWith("apt") || low.contains("apartment")) return true;
        }
        return false;
    }

    public boolean isOfficeSpace(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t != null) if ("OFFICE".equalsIgnoreCase(t.name())) return true;

        String id = s.getIdentifier();
        if (id != null) {
            String low = id.toLowerCase(Locale.ROOT);
            if (low.contains("офис") || low.contains("office")) return true;
        }
        Floor f = findFloorForSpace(s);
        return isOfficeFloor(f);
    }

    public boolean isPublicSpace(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t != null) {
            String tn = t.name();
            if ("PUBLIC".equalsIgnoreCase(tn) ||
                    "PUBLIC_AREA".equalsIgnoreCase(tn) ||
                    "COMMON".equalsIgnoreCase(tn) ||
                    "COMMUNITY".equalsIgnoreCase(tn)) return true;
        }
        String id = s.getIdentifier();
        if (id != null) {
            String low = id.toLowerCase(Locale.ROOT);
            if (low.contains("обществен") || low.contains("общ.") || low.contains("общее")) return true;
        }
        Floor f = findFloorForSpace(s);
        return isPublicFloor(f);
    }

    /* ===================== ТОПЫ НАЗВАНИЙ ===================== */

    public List<String> collectPopularApartmentRoomNames(int limit) {
        Map<String, Integer> freq = new HashMap<>();
        Map<String, String> display = new LinkedHashMap<>();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!isApartmentSpace(s)) continue;
                for (Room r : s.getRooms()) {
                    String raw = (r.getName() == null) ? "" : r.getName().trim();
                    if (raw.isEmpty()) continue;
                    String base = normalizeRoomBaseName(raw);
                    if (base.isEmpty()) continue;
                    String key = base.toLowerCase(Locale.ROOT);
                    freq.merge(key, 1, Integer::sum);
                    display.putIfAbsent(key, base);
                }
            }
        }
        return freq.entrySet().stream()
                .sorted((a, c) -> {
                    int byCount = Integer.compare(c.getValue(), a.getValue());
                    return (byCount != 0) ? byCount : display.get(a.getKey())
                            .compareToIgnoreCase(display.get(c.getKey()));
                })
                .limit(limit)
                .map(e -> display.get(e.getKey()))
                .toList();
    }

    public List<String> collectPopularPublicRoomNames(int limit) {
        Map<String, Integer> freq = new HashMap<>();
        Map<String, String> display = new LinkedHashMap<>();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!isPublicSpace(s)) continue;
                for (Room r : s.getRooms()) {
                    String raw = (r.getName() == null) ? "" : r.getName().trim();
                    if (raw.isEmpty()) continue;
                    String base = normalizePublicBaseName(raw);
                    if (base.isEmpty()) continue;
                    String key = base.toLowerCase(Locale.ROOT);
                    freq.merge(key, 1, Integer::sum);
                    display.putIfAbsent(key, base);
                }
            }
        }
        return freq.entrySet().stream()
                .sorted((a, c) -> {
                    int byCount = Integer.compare(c.getValue(), a.getValue());
                    return (byCount != 0) ? byCount : display.get(a.getKey())
                            .compareToIgnoreCase(display.get(c.getKey()));
                })
                .limit(limit)
                .map(e -> display.get(e.getKey()))
                .toList();
    }

    public List<String> collectPopularOfficeRoomNames(int limit) {
        Map<String, Integer> freq = new HashMap<>();
        Map<String, String> display = new LinkedHashMap<>();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!isOfficeSpace(s)) continue;
                for (Room r : s.getRooms()) {
                    String raw = (r.getName() == null) ? "" : r.getName().trim();
                    if (raw.isEmpty()) continue;
                    String base = normalizeOfficeBaseName(raw);
                    if (base.isEmpty()) continue;
                    String key = base.toLowerCase(Locale.ROOT);
                    freq.merge(key, 1, Integer::sum);
                    display.putIfAbsent(key, base);
                }
            }
        }
        return freq.entrySet().stream()
                .sorted((a, c) -> {
                    int byCount = Integer.compare(c.getValue(), a.getValue());
                    return (byCount != 0) ? byCount : display.get(a.getKey())
                            .compareToIgnoreCase(display.get(c.getKey()));
                })
                .limit(limit)
                .map(e -> display.get(e.getKey()))
                .toList();
    }

    /* ===================== КОПИРОВАНИЕ/ИДЕНТИФИКАТОРЫ ===================== */

    public Room createRoomCopy(Room originalRoom) {
        Room roomCopy = new Room();
        roomCopy.setId(UUID.randomUUID().hashCode());
        roomCopy.setName(originalRoom.getName());
        roomCopy.setVolume(originalRoom.getVolume());
        roomCopy.setVentilationChannels(originalRoom.getVentilationChannels());
        roomCopy.setVentilationSectionArea(originalRoom.getVentilationSectionArea());
        roomCopy.setSelected(false); // освещение сбрасываем
        try { roomCopy.setExternalWallsCount(originalRoom.getExternalWallsCount()); } catch (Throwable ignore) {}
        roomCopy.setOriginalRoomId(originalRoom.getId());
        return roomCopy;
    }

    public Space createSpaceCopyWithNewIds(Space originalSpace) {
        Space spaceCopy = new Space();
        spaceCopy.setIdentifier(originalSpace.getIdentifier());
        spaceCopy.setType(originalSpace.getType());
        for (Room originalRoom : originalSpace.getRooms()) {
            Room roomCopy = createRoomCopy(originalRoom);
            spaceCopy.addRoom(roomCopy);
        }
        return spaceCopy;
    }

    public Floor createFloorCopy(Floor originalFloor) {
        Floor floorCopy = new Floor();
        floorCopy.setNumber(originalFloor.getNumber());
        floorCopy.setType(originalFloor.getType());
        floorCopy.setName(originalFloor.getName());
        floorCopy.setSectionIndex(originalFloor.getSectionIndex());
        floorCopy.setPosition(originalFloor.getPosition());
        for (Space origSpace : originalFloor.getSpaces()) {
            Space spaceCopy = createSpaceCopyWithNewIds(origSpace);
            floorCopy.addSpace(spaceCopy);
        }
        return floorCopy;
    }

    public Floor createFloorCopyPreserve(Floor original) {
        Floor copy = new Floor();
        copy.setNumber(original.getNumber());
        copy.setType(original.getType());
        copy.setName(original.getName());
        copy.setSectionIndex(original.getSectionIndex());
        copy.setPosition(original.getPosition());
        for (Space os : original.getSpaces()) {
            copy.addSpace(createSpaceCopyPreserve(os));
        }
        return copy;
    }

    public Space createSpaceCopyPreserve(Space original) {
        Space copy = new Space();
        copy.setIdentifier(original.getIdentifier());
        copy.setType(original.getType());
        copy.setPosition(original.getPosition());
        for (Room or : original.getRooms()) {
            Room r = new Room();
            r.setId(or.getId());
            r.setName(or.getName());
            r.setVolume(or.getVolume());
            r.setVentilationChannels(or.getVentilationChannels());
            r.setVentilationSectionArea(or.getVentilationSectionArea());
            r.setSelected(or.isSelected());
            r.setMicroclimateSelected(or.isMicroclimateSelected());
            r.setRadiationSelected(or.isRadiationSelected());
            r.setOriginalRoomId(or.getOriginalRoomId());
            try { r.setExternalWallsCount(or.getExternalWallsCount()); } catch (Throwable ignore) {}
            r.setPosition(or.getPosition());
            copy.addRoom(r);
        }
        return copy;
    }

    public void updateSpaceIdentifiers(Floor floor, String newFloorDigits) {
        for (Space space : floor.getSpaces()) {
            String newId = updateIdentifier(space.getIdentifier(), newFloorDigits);
            space.setIdentifier(newId);
        }
    }

    public String updateIdentifier(String identifier, String newFloorDigits) {
        Pattern pattern = Pattern.compile("(.*?)(\\d+)-(\\d+)(.*)");
        Matcher matcher = pattern.matcher(identifier);
        if (matcher.matches()) {
            String prefix = matcher.group(1);
            String roomNum = matcher.group(3);
            String suffix = matcher.group(4);
            return prefix + newFloorDigits + "-" + roomNum + suffix;
        }
        return identifier;
    }

    public String extractDigits(String input) {
        return input.replaceAll("\\D", "");
    }

    public String generateNextFloorNumber(String currentNumber, int sectionIndex) {
        Pattern p = Pattern.compile("\\d+");
        Matcher m = p.matcher(currentNumber);
        if (m.find()) {
            int baseNum = Integer.parseInt(m.group());
            String prefix = currentNumber.substring(0, m.start());
            String suffix = currentNumber.substring(m.end());

            int maxNumberInSection = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == sectionIndex)
                    .map(Floor::getNumber)
                    .mapToInt(n -> {
                        Matcher mm = p.matcher(n);
                        return mm.find() ? Integer.parseInt(mm.group()) : Integer.MIN_VALUE;
                    })
                    .max()
                    .orElse(baseNum);

            int next = (maxNumberInSection != Integer.MIN_VALUE)
                    ? Math.max(baseNum, maxNumberInSection) + 1
                    : baseNum + 1;

            return prefix + next + suffix;
        }
        return generateUniqueNonNumericNameWithinSection(currentNumber, sectionIndex);
    }

    public String generateNextFloorNumber(String currentNumber) {
        return generateNextFloorNumber(currentNumber, -1);
    }

    public String generateUniqueNonNumericNameWithinSection(String base, int sectionIndex) {
        Pattern pattern = Pattern.compile(Pattern.quote(base) + "(?: \\(копия (\\d+)\\))?");
        int maxCopy = building.getFloors().stream()
                .filter(f -> sectionIndex < 0 || f.getSectionIndex() == sectionIndex)
                .map(Floor::getNumber)
                .map(pattern::matcher)
                .filter(Matcher::matches)
                .mapToInt(m -> m.group(1) != null ? Integer.parseInt(m.group(1)) : 0)
                .max()
                .orElse(0);

        return base + (maxCopy == 0 ? "" : " (копия " + (maxCopy + 1) + ")");
    }

    /* ===================== МИКРОКЛИМАТ: снимки и автопроставления ===================== */

    public Map<String, Boolean> saveMicroclimateSelections() {
        Map<String, Boolean> map = new HashMap<>();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    String key = f.getNumber() + "|" + s.getIdentifier() + "|" + r.getName();
                    map.put(key, r.isMicroclimateSelected());
                }
            }
        }
        return map;
    }

    public void restoreMicroclimateSelections(Map<String, Boolean> saved) {
        if (saved == null || saved.isEmpty()) return;
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    String key = f.getNumber() + "|" + s.getIdentifier() + "|" + r.getName();
                    Boolean v = saved.get(key);
                    if (v != null) r.setMicroclimateSelected(v);
                }
            }
        }
    }

    /** Для офисов (по типу помещения ИЛИ по типу этажа) ставим МК всем, кроме санузлов */
    public void applyMicroDefaultsForOfficeSpace(Space s) {
        if (s == null) return;
        if (!isOfficeSpace(s)) return;
        for (Room r : s.getRooms()) {
            String n = (r.getName() == null) ? "" : r.getName();
            r.setMicroclimateSelected(!looksLikeSanitary(n));
        }
    }
}
