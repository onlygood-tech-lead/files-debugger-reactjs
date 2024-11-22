/**
 * Normalizes a given string by removing leading and trailing whitespace,
 * converting to lowercase, and removing all non-alphanumeric characters except spaces.
 *
 * @param phrase - The input string to be normalized.
 * @returns The normalized string.
 */
export function normalizeString(phrase: string): string {
  return phrase
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, "");
}

/**
 * Searches for a normalized version of the input phrase within a given array of strings.
 * The search is case-insensitive and ignores special characters and extra whitespace.
 *
 * @param phrase - The string to search for in the array.
 * @param array - The array of strings to search within.
 * @returns The normalized matching string if found, or null if not found.
 */
export function matchNormalizedPhrase(
  phrase: string,
  array: string[],
): string | null {
  const normalizedPhrase = normalizeString(phrase);

  for (let i = 0; i < array.length; i++) {
    if (normalizeString(array[i]) === normalizedPhrase) {
      return array[i]; // Return the original string from the input array
    }
  }

  return null; // Return null if no match is found
}

/**
 * Retrieves a value from a nested object using a key path.
 * @param obj The object to retrieve the value from.
 * @param keyPath An array of keys representing the path to the desired value.
 * @returns The value at the specified key path, or null if not found or in case of error.
 */
export function getNestedValue<T = any>(
  obj: Record<string, any>,
  keyPath: string[],
): T | null {
  try {
    // Validate input parameters
    if (obj === null || typeof obj !== "object") {
      throw new Error("Invalid input: obj must be a non-null object");
    }

    if (!Array.isArray(keyPath) || keyPath.length === 0) {
      throw new Error(
        "Invalid input: keyPath must be a non-empty array of strings",
      );
    }

    if (!keyPath.every((key) => typeof key === "string")) {
      throw new Error("Invalid input: all elements in keyPath must be strings");
    }

    // Traverse the object using the key path
    let current: any = obj;
    for (const key of keyPath) {
      if (current === null || typeof current !== "object") {
        return null; // Early return if we can't traverse further
      }
      if (!(key in current)) {
        return null; // Key not found
      }
      current = current[key];
    }

    return current as T;
  } catch (error) {
    console.error("Error in getNestedValue(): ", error);
    return null;
  }
}
