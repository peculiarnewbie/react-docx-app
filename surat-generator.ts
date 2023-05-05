import * as fs from 'fs';
import { saveAs } from 'file-saver';
import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TabStopPosition,
  TabStopType,
  TextRun,
  Header,
  ImageRun,
  ShadingType,
  UnderlineType,
  convertInchesToTwip,
  BorderStyle,
  HorizontalPositionRelativeFrom,
} from 'docx';
const PHONE_NUMBER = '07534563401';
const PROFILE_URL = 'https://www.linkedin.com/in/dolan1';
const EMAIL = 'docx@docx.com';

export class DocumentCreator {
  public nomor;
  public nama;
  public dokter;
  public berat;
  public tinggi;

  constructor(nomor, nama, dokter, berat, tinggi) {
    this.nomor = nomor;
    this.nama = nama;
    this.dokter = dokter;
    this.berat = berat;
    this.tinggi = tinggi;
  }

  // tslint:disable-next-line: typedef
  async create() {
    const kemenkes = await fetch(
      'https://raw.githubusercontent.com/peculiarnewbie/react-docx-app/main/images/logo-kesehatan.png'
    ).then((r) => r.blob());

    const mdjamil = await fetch(
      'https://raw.githubusercontent.com/peculiarnewbie/react-docx-app/main/images/rsupmdjamil.png'
    ).then((r) => r.blob());

    const doc = new Document({
      styles: {
        default: {
          heading1: {
            run: {
              size: 28,
              bold: true,
              color: '000000',
              font: 'Arial',
            },
            paragraph: {
              spacing: {
                after: 0,
                before: 0,
              },
            },
          },
          heading2: {
            run: {
              size: 26,
              bold: true,
              color: '000000',
              font: 'Arial',
            },
            paragraph: {
              spacing: {
                before: 0,
                after: 0,
              },
            },
          },
          heading3: {
            run: {
              size: 24,
              color: '000000',
              font: 'Arial',
            },
            paragraph: {
              spacing: {
                before: 0,
                after: 0,
              },
            },
          },
          listParagraph: {
            run: {
              color: '#FF0000',
            },
          },
        },
        paragraphStyles: [
          {
            id: 'aside',
            name: 'Aside',
            basedOn: 'Normal',
            next: 'Normal',
            run: {
              color: '999999',
              italics: true,
            },
            paragraph: {
              indent: {
                left: convertInchesToTwip(0.5),
              },
              spacing: {
                line: 276,
              },
            },
          },
          {
            id: 'wellSpaced',
            name: 'Well Spaced',
            basedOn: 'Normal',
            quickFormat: true,
            paragraph: {
              spacing: {
                line: 276,
                before: 20 * 72 * 0.1,
                after: 20 * 72 * 0.05,
              },
            },
          },
          {
            id: 'headerLine',
            name: 'Header Line',
            basedOn: 'Normal',
            quickFormat: true,
            paragraph: {
              indent: {
                left: -500,
                right: -500,
              },
            },
          },
        ],
      },
      sections: [
        {
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: kemenkes,
                      transformation: {
                        width: 80,
                        height: 80,
                      },
                      floating: {
                        horizontalPosition: {
                          offset: 600000, // relative: HorizontalPositionRelativeFrom.PAGE by default
                        },
                        verticalPosition: {
                          offset: 480000, // relative: VerticalPositionRelativeFrom.PAGE by default
                        },
                      },
                    }),
                    new ImageRun({
                      data: mdjamil,
                      transformation: {
                        width: 80,
                        height: 80,
                      },
                      floating: {
                        horizontalPosition: {
                          relative: HorizontalPositionRelativeFrom.RIGHT_MARGIN,
                          offset: -600000, // relative: HorizontalPositionRelativeFrom.PAGE by default
                        },
                        verticalPosition: {
                          offset: 480000, // relative: VerticalPositionRelativeFrom.PAGE by default
                        },
                      },
                    }),
                    new Paragraph({
                      text: 'KEMENTRIAN KESEHATAN REPUBLIK INDONESIA',
                      heading: HeadingLevel.HEADING_1,
                      alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                      text: 'DIREKTORAT JENDERAL PELAYANAN KESEHATAN',
                      heading: HeadingLevel.HEADING_2,
                      alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                      text: 'RUMAH SAKIT UMUM PUSAT DR. M. DJAMIL PADANG',
                      heading: HeadingLevel.HEADING_3,
                      alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                      text: 'Jalan Perintis Kemerdekaan Padang - 25127',
                      alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                      text: 'Phone: (0751) 32371, 810253, 810254 Fax : (0751) 32371',
                      alignment: AlignmentType.CENTER,
                    }),
                    new Paragraph({
                      text: 'Website : www.rsdjamil.co.id, email : rsupdjamil@yahoo.com',
                      alignment: AlignmentType.CENTER,
                      style: 'headerLine',
                      border: {
                        bottom: {
                          color: 'auto',
                          style: BorderStyle.THIN_THICK_THIN_SMALL_GAP,
                          size: 12,
                        },
                      },
                    }),
                  ],
                }),
              ],
            }),
          },
          children: [
            this.createHeaderSurat(this.nomor),
            new Paragraph({
              text: this.nama,
              heading: HeadingLevel.TITLE,
            }),
            new Paragraph({
              text: `Berat: ${this.berat}kg`,
              alignment: AlignmentType.LEFT,
            }),
            new Paragraph({
              text: `Tinggi: ${this.tinggi}cm`,
              alignment: AlignmentType.LEFT,
            }),

            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun(`Dokter: ${this.dokter}`),
                new TextRun({
                  text: '________________',
                  break: 5,
                }),
              ],
            }),
          ],
        },
      ],
    });

    Packer.toBlob(doc).then((blob) => {
      console.log(blob);
      saveAs(blob, `${this.nomor}.${this.nama}.docx`);
      console.log('Document created successfully');
    });
  }

  public createHeaderSurat(nomor: string): Paragraph {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun(`Surat no:  ${nomor}`)],
    });
  }

  public createContactInfo(
    phoneNumber: string,
    profileUrl: string,
    email: string
  ): Paragraph {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun(
          `Mobile: ${phoneNumber} | LinkedIn: ${profileUrl} | Email: ${email}`
        ),
        new TextRun({
          text: 'Address: 58 Elm Avenue, Kent ME4 6ER, UK',
          break: 1,
        }),
      ],
    });
  }

  public createHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_1,
      thematicBreak: true,
    });
  }

  public createSubHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_2,
    });
  }

  public createInstitutionHeader(
    institutionName: string,
    dateText: string
  ): Paragraph {
    return new Paragraph({
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: TabStopPosition.MAX,
        },
      ],
      children: [
        new TextRun({
          text: institutionName,
          bold: true,
        }),
        new TextRun({
          text: `\t${dateText}`,
          bold: true,
        }),
      ],
    });
  }

  public createRoleText(roleText: string): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: roleText,
          italics: true,
        }),
      ],
    });
  }

  public createBullet(text: string): Paragraph {
    return new Paragraph({
      text: text,
      bullet: {
        level: 0,
      },
    });
  }

  // tslint:disable-next-line:no-any
  public createSkillList(skills: any[]): Paragraph {
    return new Paragraph({
      children: [
        new TextRun(skills.map((skill) => skill.name).join(', ') + '.'),
      ],
    });
  }

  // tslint:disable-next-line:no-any
  public createAchivementsList(achivements: any[]): Paragraph[] {
    return achivements.map(
      (achievement) =>
        new Paragraph({
          text: achievement.name,
          bullet: {
            level: 0,
          },
        })
    );
  }

  public createInterests(interests: string): Paragraph {
    return new Paragraph({
      children: [new TextRun(interests)],
    });
  }

  public splitParagraphIntoBullets(text: string): string[] {
    return text.split('\n\n');
  }

  // tslint:disable-next-line:no-any
  public createPositionDateText(
    startDate: any,
    endDate: any,
    isCurrent: boolean
  ): string {
    const startDateText =
      this.getMonthFromInt(startDate.month) + '. ' + startDate.year;
    const endDateText = isCurrent
      ? 'Present'
      : `${this.getMonthFromInt(endDate.month)}. ${endDate.year}`;

    return `${startDateText} - ${endDateText}`;
  }

  public getMonthFromInt(value: number): string {
    switch (value) {
      case 1:
        return 'Jan';
      case 2:
        return 'Feb';
      case 3:
        return 'Mar';
      case 4:
        return 'Apr';
      case 5:
        return 'May';
      case 6:
        return 'Jun';
      case 7:
        return 'Jul';
      case 8:
        return 'Aug';
      case 9:
        return 'Sept';
      case 10:
        return 'Oct';
      case 11:
        return 'Nov';
      case 12:
        return 'Dec';
      default:
        return 'N/A';
    }
  }
}
