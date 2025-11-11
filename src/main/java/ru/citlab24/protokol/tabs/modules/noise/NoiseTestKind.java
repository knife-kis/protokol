package ru.citlab24.protokol.tabs.modules.noise;

/** Типы испытаний для шумов (лифт/ИТО/авто/площадка). */
public enum NoiseTestKind {
    LIFT_DAY,     // Лифт — день
    LIFT_NIGHT,   // Лифт — ночь

    ITO_NONRES,   // «шим неж ИТО»
    ITO_RES,      // «шум жил ИТО»

    AUTO_DAY,     // «шум авто день»
    AUTO_NIGHT,   // «шум авто ночь»

    SITE          // «шум площадка»
}
