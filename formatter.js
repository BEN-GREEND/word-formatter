Office.onReady(function(info) {
    // כשהדף נטען, הוסף את הפונקציה לכפתור
    document.getElementById('formatButton').onclick = formatText;
});

function distributeSpaces(shortLine, targetLength) {
    if (!shortLine.trim()) return shortLine;
    
    const words = shortLine.trim().split(/\s+/);
    if (words.length <= 1) return shortLine.padEnd(targetLength);
    
    const currentLength = shortLine.trim().length;
    const spacesToAdd = targetLength - currentLength;
    const gaps = words.length - 1;
    const spacesPerGap = Math.floor(spacesToAdd / gaps);
    const extraSpaces = spacesToAdd % gaps;
    
    return words.reduce((result, word, index) => {
        if (index === words.length - 1) return result + word;
        const extraSpace = index < extraSpaces ? 1 : 0;
        const spaces = ' '.repeat(1 + spacesPerGap + extraSpace);
        return result + word + spaces;
    }, '');
}

async function formatText() {
    try {
        await Word.run(async (context) => {
            // קבלת הטקסט שנבחר
            const selection = context.document.getSelection();
            selection.load("text");
            await context.sync();

            // פיצול הטקסט לשורות
            const lines = selection.text.split('\n');
            const processedLines = [];

            // עיבוד כל שורה
            for (let line of lines) {
                // פיצול השורה לטורים (לפי טאב)
                const columns = line.split('\t');
                if (columns.length < 2) continue;  // דלג על שורות ללא שני טורים

                // חישוב האורך המקסימלי
                const col1Length = columns[0].trim().length;
                const col2Length = columns[1].trim().length;

                // עיבוד הטורים
                if (col1Length > col2Length) {
                    columns[1] = distributeSpaces(columns[1], col1Length);
                } else {
                    columns[0] = distributeSpaces(columns[0], col2Length);
                }

                // חיבור הטורים בחזרה
                processedLines.push(columns.join('\t'));
            }

            // החלפת הטקסט המקורי
            const newText = processedLines.join('\n');
            selection.insertText(newText, Word.InsertLocation.replace);

            await context.sync();
        });
    } catch (error) {
        console.error('Error:', error);
    }
}
