console.log('Client-side code running');

function Generate() {
    const doc = new docx.Document({
        sections: [{
            children: [
                new docx.Paragraph({
                    children: [new docx.TextRun("Lorem Ipsum Foo Bar"), new docx.TextRun("Hello World")],
                }),
            ],
        }]
    });
}
