import {
    trimToLatestReply,
    trimSignature,
    capBlankRuns,
    trimOuterBlankLines
} from "./text-utils.js";

async function getBody(item) {
    return new Promise((resolve, reject) => {

        item.body.getAsync(
            Office.CoercionType.Text,
            result => {

                if (
                    result.status ===
                    Office.AsyncResultStatus.Succeeded
                ) {
                    resolve(result.value || "");
                }
                else {
                    reject(result.error);
                }
            }
        );
    });
}

function formatAddress(person) {

    if (!person) {
        return "";
    }

    const name =
        person.displayName || "";

    const email =
        person.emailAddress || "";

    if (!email) {
        return name;
    }

    return `${name} <${email}>`;
}

function joinRecipients(list = []) {

    return list
        .map(formatAddress)
        .join("; ");
}

async function copyEmailDetails(event) {

    try {

        const item =
            Office.context.mailbox.item;

        let body =
            await getBody(item);

        body =
            trimToLatestReply(body);

        body =
            trimSignature(
                body,
                item?.from?.displayName || ""
            );

        body =
            capBlankRuns(body, 1);

        body =
            trimOuterBlankLines(body);

        const output = [
            `From: ${formatAddress(item.from)}`,
            `To: ${joinRecipients(item.to)}`,
            item.cc?.length
                ? `Cc: ${joinRecipients(item.cc)}`
                : null,
            `Subject: ${item.subject || ""}`,
            "",
            body
        ]
        .filter(Boolean)
        .join("\r\n");

        await navigator.clipboard.writeText(
            output
        );

        item.notificationMessages.replaceAsync(
            "copied",
            {
                type:
                    Office.MailboxEnums
                    .ItemNotificationMessageType
                    .InformationalMessage,

                message:
                    "Email copied to clipboard.",

                persistent: false
            }
        );
    }
    catch (err) {

        Office.context.mailbox.item
            .notificationMessages
            .replaceAsync(
                "error",
                {
                    type:
                        Office.MailboxEnums
                        .ItemNotificationMessageType
                        .ErrorMessage,

                    message:
                        err.message ||
                        "Unable to copy email."
                }
            );
    }

    event.completed();
}

Office.onReady(() => {

    Office.actions.associate(
        "copyEmailDetails",
        copyEmailDetails
    );
});