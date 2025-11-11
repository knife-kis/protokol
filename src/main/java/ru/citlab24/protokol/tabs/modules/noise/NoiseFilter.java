package ru.citlab24.protokol.tabs.modules.noise;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import java.util.*;
import java.util.stream.Collectors;

/** Инкапсулирует глобальную фильтрацию по источникам шумов. */
public final class NoiseFilter {

    private final Building building;
    private final Map<String, DatabaseManager.NoiseValue> byKey;
    private final Set<String> active; // активные источники (подписи кнопок)

    public NoiseFilter(Building building,
                       Map<String, DatabaseManager.NoiseValue> byKey,
                       Set<String> active) {
        this.building = (building != null) ? building : new Building();
        this.byKey = (byKey != null) ? byKey : Collections.emptyMap();
        this.active = (active != null) ? new LinkedHashSet<>(active) : Collections.emptySet();
    }

    public boolean isActive() { return !active.isEmpty(); }

    /** Секции, где есть хотя бы одна комната, проходящая фильтр. */
    public List<Section> filterSections() {
        List<Section> all = building.getSections();
        if (!isActive()) return all;
        List<Section> out = new ArrayList<>();
        for (int secIdx = 0; secIdx < all.size(); secIdx++) {
            if (hasSectionMatch(secIdx)) out.add(all.get(secIdx));
        }
        return out;
    }

    /** Этажи выбранной секции. */
    public List<Floor> filterFloors(int secIdx) {
        List<Floor> floors = building.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .sorted(Comparator.comparingInt(Floor::getPosition))
                .collect(Collectors.toList());
        if (!isActive()) return floors;
        return floors.stream().filter(f -> hasFloorMatch(secIdx, f)).collect(Collectors.toList());
    }

    /** Помещения на этаже. */
    public List<Space> filterSpaces(int secIdx, Floor f) {
        if (f == null) return Collections.emptyList();
        List<Space> spaces = new ArrayList<>(f.getSpaces());
        spaces.sort(Comparator.comparingInt(Space::getPosition));
        if (!isActive()) return spaces;
        return spaces.stream().filter(s -> hasSpaceMatch(secIdx, f, s)).collect(Collectors.toList());
    }

    /** Комнаты в помещении. */
    public List<Room> filterRooms(int secIdx, Floor f, Space s) {
        if (s == null) return Collections.emptyList();
        List<Room> rooms = new ArrayList<>(s.getRooms());
        rooms.sort(Comparator.comparingInt(Room::getPosition));
        if (!isActive()) return rooms;
        return rooms.stream().filter(r -> roomMatchesFilter(secIdx, f, s, r)).collect(Collectors.toList());
    }

    /* ===== внутренняя логика соответствий ===== */

    private boolean hasSectionMatch(int secIdx) {
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() == secIdx && hasFloorMatch(secIdx, f)) return true;
        }
        return false;
    }

    private boolean hasFloorMatch(int secIdx, Floor f) {
        for (Space s : f.getSpaces()) if (hasSpaceMatch(secIdx, f, s)) return true;
        return false;
    }

    private boolean hasSpaceMatch(int secIdx, Floor f, Space s) {
        for (Room r : s.getRooms()) if (roomMatchesFilter(secIdx, f, s, r)) return true;
        return false;
    }

    private boolean roomMatchesFilter(int secIdx, Floor f, Space s, Room r) {
        if (!isActive()) return true;
        String key = makeKey(secIdx, f, s, r);
        DatabaseManager.NoiseValue nv = byKey.get(key);
        Set<String> sources = (nv == null) ? Collections.emptySet() : getNvSources(nv);
        for (String a : active) if (sources.contains(a)) return true;
        return false;
    }

    private static String makeKey(int secIdx, Floor f, Space s, Room r) {
        String floorNum = (f == null || f.getNumber() == null) ? "" : f.getNumber().trim();
        String spaceId  = (s == null || s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
        String roomName = (r == null || r.getName() == null) ? "" : r.getName().trim();
        return secIdx + "|" + floorNum + "|" + spaceId + "|" + roomName;
    }

    /** Маппинг полей NoiseValue -> подписи кнопок. */
    private static Set<String> getNvSources(DatabaseManager.NoiseValue nv) {
        Set<String> s = new LinkedHashSet<>();
        if (nv.lift)        s.add("Лифт");
        if (nv.vent)        s.add("Вент");
        if (nv.heatCurtain) s.add("Завеса");
        if (nv.itp)         s.add("ИТП");
        if (nv.pns)         s.add("ПНС");
        if (nv.electrical)  s.add("Э/Щ");
        if (nv.autoSrc)     s.add("Авто");
        if (nv.zum)         s.add("Зум");
        return s;
    }
}
