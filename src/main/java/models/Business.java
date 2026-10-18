package models;

import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * A business from the leads sheet, with all of its rows put together.
 *
 * @param name    the business name (שם עסק)
 * @param details its filled-in details, label to value (e.g. "טלפון" to "050-…"), in the sheet's column order
 * @param wanted  the professions it asked for and how many workers of each, in the sheet's row order
 */
public record Business(String name, Map<String, String> details, Map<String, Integer> wanted) {
    public Business {
        // Copies that keep the order and cannot be changed.
        details = Collections.unmodifiableMap(new LinkedHashMap<>(details));
        wanted = Collections.unmodifiableMap(new LinkedHashMap<>(wanted));
    }

    /** The total number of workers the business asked for. */
    public int totalWanted() {
        return wanted.values().stream().mapToInt(Integer::intValue).sum();
    }
}
