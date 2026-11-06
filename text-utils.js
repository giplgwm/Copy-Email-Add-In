export function trimToLatestReply(text) {

    const lines =
        text.split(/\r?\n/);

    for (
        let i = 0;
        i < lines.length;
        i++
    ) {

        const line =
            lines[i].trim();

        if (
            /^On .* wrote:$/i.test(line) ||
            /^From:/i.test(line) ||
            /Original Message/i.test(line)
        ) {
            return lines
                .slice(0, i)
                .join("\n");
        }
    }

    return text;
}

export function trimSignature(
    text,
    senderName
) {

    const lines =
        text.split(/\r?\n/);

    const signatures = [

        "thanks",
        "thank you",
        "best",
        "best regards",
        "kind regards",
        "regards",
        "cheers",
        "sincerely",
        "v/r",

        senderName.toLowerCase(),

        "sent from my iphone",
        "get outlook for ios",
        "get outlook for android",
        "confidentiality notice"
    ];

    for (
        let i = 0;
        i < lines.length;
        i++
    ) {

        const lower =
            lines[i]
                .trim()
                .toLowerCase();

        if (
            signatures.some(
                sig =>
                    lower === sig
            )
        ) {
            return lines
                .slice(0, i)
                .join("\n");
        }
    }

    return text;
}

export function capBlankRuns(
    text,
    max = 1
) {

    const lines =
        text.split(/\r?\n/);

    const output = [];

    let blankCount = 0;

    for (const line of lines) {

        if (
            line.trim() === ""
        ) {

            blankCount++;

            if (
                blankCount > max
            ) {
                continue;
            }
        }
        else {
            blankCount = 0;
        }

        output.push(line);
    }

    return output.join("\n");
}

export function trimOuterBlankLines(
    text
) {

    return text.replace(
        /^\s+|\s+$/g,
        ""
    );
}