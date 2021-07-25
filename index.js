console.log('Server-side code running');

const { HorizontalPositionAlign } = require('docx');
const docx = require('docx')
const { Packer, HeadingLevel } = docx;
const express = require('express');
const bodyParser = require('body-parser')
const fs = require('fs');
const path = require('path');
let converter = require('docx-pdf');
const unoconv = require('awesome-unoconv');

// DATA
let To, From, Date, Customer, EquipmentMain, Materials, PurposeOfArrival, Equipment, Product, Temperature, StatusOfMaterials, DataOnMaterials
let Result, Conclusion
let Heating, HeatingPressure, Molding, Pressure, TemperatureTop, TemperatureLow, Sealing, VentilationDelay, WeldingTemperature, GasMixture, Format, Cell

function Generate() {

    const MainTable = new docx.Table({
        columnWidths: [4505, 4505],
        rows: [
            // 1 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.SINGLE,
                                size: 10
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "To/Кому:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.SINGLE,
                                size: 10
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${To}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
            // 2 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "From/От кого:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${From}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
            // 3 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Date/Дата:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Date}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
            // 4 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Customer/Клиент:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Customer}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
            // 5 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Equipment/Оборудование:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${EquipmentMain}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
            // 6 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.SINGLE,
                                size: 10
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Materials/Материалы:",
                            style: "main-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        borders: {
                            top: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            bottom: {
                                style: docx.BorderStyle.SINGLE,
                                size: 10
                            },
                            left: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            },
                            right: {
                                style: docx.BorderStyle.NONE,
                                size: 1
                            }
                        },
                        width: {
                            size: 4505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Materials}`,
                            style: "main-table-user-text"
                        })],
                    }),
                ],
            }),
        ],
    });

    const SecondTable = new docx.Table({
        columnWidths: [6505, 6505],
        rows: [
            // 1 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Формование",
                            style: "second-table-text",
                            alignment: docx.AlignmentType.CENTER,
                        })],
                        columnSpan: 2,
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Сварка",
                            style: "second-table-text",
                            alignment: docx.AlignmentType.CENTER,
                        })],
                        columnSpan: 2,
                    }),
                ],
            }),
            // 2 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Нагрев",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Heating}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Запечатывание",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Sealing}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
            // 3 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Давление нагрева",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${HeatingPressure}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Задерж. вентиляции",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${VentilationDelay}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
            // 4 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Формование",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Molding}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Температура",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${WeldingTemperature}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
            // 5 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Форм. уст. давление",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Pressure}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Газовая смесь",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${GasMixture}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
            // 6 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Температура (верх)",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${TemperatureTop}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Параметры формы",
                            style: "second-table-text",
                            alignment: docx.AlignmentType.CENTER,
                        })],
                        columnSpan: 2,
                    }),
                ],
            }),
            // 7 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Температура (низ)",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${TemperatureLow}`,
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Формат",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Format}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
            // 7 LINE
            new docx.TableRow({
                children: [
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: "Ячейка",
                            style: "second-table-text"
                        })],
                    }),
                    new docx.TableCell({
                        width: {
                            size: 6505,
                            type: docx.WidthType.DXA,
                        },
                        children: [new docx.Paragraph({
                            text: `${Cell}`,
                            style: "second-table-text"
                        })],
                    }),
                ],
            }),
        ],
        width: {
            size: docx.convertInchesToTwip(7.3),
            type: docx.WidthType.DXA,
        },
        alignment: docx.AlignmentType.CENTER,
    });
          

    const doc = new docx.Document({
        styles: {
            default: {
                heading1: {
                    run: {
                        size: 48,
                        color: "008A97",
                        bold: true,
                        font: {
                            name: "Abadi",
                        }
                    },
                },
                heading2: {
                    run: {
                        size: 24,
                        bold: true,
                        font: {
                            name: "Arial",
                        }
                    },
                    paragraph: {
                        indent: {
                            left: docx.convertInchesToTwip(0.55),
                        },
                        spacing: {
                            after: 100,
                        },
                    },
                }
            },
            paragraphStyles: [
                {
                    id: "main-table-text",
                    name: "MainTableText",
                    run: {
                        size: 22,
                        bold: true,
                        font: {
                            name: "Arial",
                        }
                    },
                    paragraph: {
                        indent: {
                            left: docx.convertInchesToTwip(0.55),
                        },
                        spacing: {
                            after: 100,
                        },
                    },
                },
                {
                    id: "main-table-user-text",
                    name: "MainTableUserText",
                    run: {
                        size: 22,
                        font: {
                            name: "Arial",
                        }
                    },
                    paragraph: {
                        indent: {
                            left: docx.convertInchesToTwip(0.55),
                        },
                        spacing: {
                            after: 100,
                        },
                    },
                },
                {
                    id: "text",
                    name: "Text",
                    run: {
                        size: 24,
                        font: {
                            name: "Arial",
                        }
                    },
                    // paragraph: {
                    //     indent: {
                    //         left: docx.convertInchesToTwip(0.55),
                    //     },
                    //     spacing: {
                    //         after: 100,
                    //     },
                    // },
                },
                {
                    id: "textH2",
                    name: "TextH2",
                    run: {
                        size: 24,
                        font: {
                            name: "Arial",
                        }
                    },
                    paragraph: {
                        indent: {
                            left: docx.convertInchesToTwip(0.55),
                        },
                        spacing: {
                            after: 100,
                        },
                    },
                },
                {
                    id: "second-table-text",
                    name: "Second Table Text",
                    run: {
                        size: 24,
                        font: {
                            name: "Arial",
                        }
                    },
                    paragraph: {
                        indent: {
                            left: docx.convertInchesToTwip(0.1),
                        },
                        spacing: {
                            after: 150,
                        },
                    },
                },
            ],
        },
        sections: [{
            headers: {
                default: new docx.Header({
                    children: [ 
                        new docx.Paragraph({
                            children: [
                                new docx.ImageRun({
                                    data: fs.readFileSync("Logo1.png"),
                                    transformation: {
                                        width: 202,
                                        height: 44,
                                    },
                                    // floating: {
                                    //     horizontalPosition: {
                                    //         offset: 2014400,
                                    //     },
                                    //     verticalPosition: {
                                    //         offset: 0,
                                    //     },
                                    // },
                                }),
                                new docx.ImageRun({
                                    data: fs.readFileSync("Logo2.png"),
                                    transformation: {
                                        width: 50,
                                        height: 42,
                                    },
                                    // floating: {
                                    //     horizontalPosition: {
                                    //         offset: 2014400,
                                    //     },
                                    //     verticalPosition: {
                                    //         offset: 0,
                                    //     },
                                    // },
                                }),
                            ],
                        }),
                    ],
                }),
            },
            footers: {
                default: new docx.Footer({
                    children: [new docx.Paragraph("Logo doesn't work in header and footer")],
                }),
            },
            children: [
                new docx.Paragraph({
                    text: "CIS Service & Application Report",
                    heading: HeadingLevel.HEADING_1,
                    spacing: {
                        after: 200,
                    },
                }),
                MainTable,
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: "Цель приезда",
                    heading: HeadingLevel.HEADING_2,
                }),
                new docx.Paragraph({
                    text: `${PurposeOfArrival}`,
                    style: "text"
                }),
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: "Условия проведения работ",
                    heading: HeadingLevel.HEADING_2,
                }),
                new docx.Paragraph({
                    text: `•  Оборудование: ${Equipment}`,
                    style: "textH2"
                }),
                new docx.Paragraph({
                    text: `•  Продукт: ${Product}`,
                    style: "textH2"
                }),
                new docx.Paragraph({
                    text: `•  Температура в цеху: ${Temperature}℃`,
                    style: "textH2"
                }),
                new docx.Paragraph({
                    text: `•  Состояние упаковочных материалов: ${StatusOfMaterials}`,
                    style: "textH2"
                }),
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: `Данные по материалам\n ${DataOnMaterials}`,
                    style: "textH2"
                }),
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: "Настройки оборудования:",
                    style: "textH2"
                }),
                SecondTable,
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: "Текущий результат",
                    heading: HeadingLevel.HEADING_2,
                }),
                new docx.Paragraph({
                    text: `${Result}`,
                    style: "text"
                }),
                new docx.Paragraph("\n"),
                new docx.Paragraph({
                    text: "Выводы",
                    heading: HeadingLevel.HEADING_2,
                }),
                new docx.Paragraph({
                    text: `${Conclusion}`,
                    style: "text"
                }),
                // new docx.Paragraph({
                //     children: [
                //         new docx.ImageRun({
                //             data: fs.readFileSync("Logo1.png"),
                //             transformation: {
                //                 width: 202,
                //                 height: 44,
                //             },
                //             // floating: {
                //             //     horizontalPosition: {
                //             //         offset: 2014400,
                //             //     },
                //             //     verticalPosition: {
                //             //         offset: 0,
                //             //     },
                //             // },
                //         }),
                //     ],
                // }),
            ],
        }]
    });

    Packer.toBuffer(doc).then((buffer) => {
        fs.writeFileSync("Doc.docx", buffer)
    })
}

const app = express()

// создаем парсер для данных application/x-www-form-urlencoded
const urlencodedParser = bodyParser.urlencoded({
    extended: false,
})

// serve files from the public directory
app.use(express.static('public'))

// start the express web server listening on 8080
app.listen(8080, () => {
  console.log('listening on 8080')
});

// serve the homepage
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
});

app.post('/generate.html', urlencodedParser, function (
    request,
    response) {
    if (!request.body) return response.sendStatus(400)
    console.log(request.body)
    response.send(
      "Form has been genereated succesfully"
    )

    To = request.body.To
    From = request.body.From
    Date = request.body.Date
    Customer = request.body.Customer
    EquipmentMain = request.body.EquipmentMain
    Materials = request.body.Materials
    
    PurposeOfArrival = request.body.PurposeOfArrival
    Equipment = request.body.Equipment
    Product = request.body.Product
    Temperature = request.body.Temperature
    StatusOfMaterials = request.body.StatusOfMaterials
    DataOnMaterials = request.body.DataOnMaterials

    Heating = request.body.Heating
    HeatingPressure = request.body.HeatingPressure
    Molding = request.body.Molding
    Pressure = request.body.Pressure
    TemperatureTop = request.body.TemperatureTop
    TemperatureLow = request.body.TemperatureLow
    Sealing = request.body.Sealing
    VentilationDelay = request.body.VentilationDelay
    WeldingTemperature = request.body.WeldingTemperature
    GasMixture = request.body.GasMixture
    Format = request.body.Format
    Cell = request.body.Cell

    Result = request.body.Result
    Conclusion = request.body.Conclusion
    
    Generate()

    console.log("Lol")

    converter('Doc.docx', 'output.pdf', (err, result) => {
        if (err) {
            console.log("Converting Doc to PDF failed", err);
        }
        console.log("Converting Doc to PDF succesfull", result);
    });
})