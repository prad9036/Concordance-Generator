awk '
/^#/ || NF == 0 {print; next}
{
    gsub(/^[[:space:]]+|[[:space:]]+$/, "", $0)
    key=tolower($0)
    if (!seen[key]++) print
}
' stop_word_list.txt > stop_word_list.tmp && mv stop_word_list.tmp stop_word_list.txt
