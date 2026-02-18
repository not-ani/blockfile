pub(crate) const MAX_QUERY_CHARS: usize = 512;

pub(crate) fn normalize_for_search(text: &str) -> String {
    let mut normalized = String::with_capacity(text.len());
    let mut previous_space = false;
    for character in text.chars() {
        if character.is_alphanumeric() {
            previous_space = false;
            for lower in character.to_lowercase() {
                normalized.push(lower);
            }
        } else if !previous_space {
            normalized.push(' ');
            previous_space = true;
        }
    }
    normalized.trim().to_string()
}
