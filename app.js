const DEBUG = true;

let g_xlsxData = [];
let g_members = {};
let g_monthInReports = [];

let logger;

$(document).ready(function () {
  const MSG_NUMBER_OF_FILES = "You need to upload 6 files for score analysis.";
  const KEY = {
    "LAST_NAME": "チャプター► サマリーPALMSレポート",
    "FIRST_NAME": "__EMPTY",
    "ATTENDANCE": "__EMPTY_1",
    "ABSENT": "__EMPTY_2",
    "LATE": "__EMPTY_4",
    "REFERRAL": "__EMPTY_8",
    "REFERRAL2": "__EMPTY_9",
    "VISITOR": "__EMPTY_13",
    "TYFCB": "__EMPTY_15",
    "CEU": "__EMPTY_16",
    "RECOMMENDATION": "__EMPTY_17",
    "REPORT_DATE": "__EMPTY_9",
    "REPORT_DATE_NICE": "REPORT_DATE_NICE",

    "LAST_NAME_NICE": "LAST_NAME_NICE",
    "FIRST_NAME_NICE": "FIRST_NAME_NICE",
    "ATTENDANCE_NICE": "ATTENDANCE_NICE",
    "ABSENT_NICE": "ABSENT_NICE",
    "LATE_NICE": "LATE_NICE",
    "REFERRAL_NICE": "REFERRAL_NICE",
    "REFERRAL2_NICE": "REFERRAL2_NICE",
    "REFERRAL_FINAL": "REFERRAL_FINAL",
    "VISITOR_NICE": "VISITOR_NICE",
    "TYFCB_NICE": "TYFCB_NICE",
    "CEU_NICE": "CEU_NICE",
    "RECOMMENDATION_NICE": "RECOMMENDATION_NICE",

    "ABSENT_SCORE": "ABSENT_SCORE",
    "LATE_SCORE": "LATE_SCORE",
    "REFERRAL_SCORE": "REFERRAL_SCORE",
    "VISITOR_SCORE": "VISITOR_SCORE",
    "TYFCB_SCORE": "TYFCB_SCORE",
    "CEU_SCORE": "CEU_SCORE",
    "RECOMMENDATION_SCORE": "RECOMMENDATION_SCORE",
    "TOTAL_SCORE": "TOTAL_SCORE",
  };
  logger = new Logger(DEBUG);

  $(".file-multiple").change(function (event) {
    reset();

    let files = event.target.files;
    if (files.length != 6) {
      showMessage(MSG_NUMBER_OF_FILES);
      return;
    }

    for (let i = 0, len = files.length; i < len; i++) {
      readFile(files[i]);
    }

    // List all file name and size in screen
    renderFileNames(files);

    // Load data in delay time
    setTimeout(function () {
      handleFileInput();
      updateExportButtonState(true);
    }, 1000);
  })

  $(".export-score").click(function () {
    exportXLSX();
  });

  /**
   * Render list files with name and size
   * @param  {array} files list of files
   */
  function renderFileNames(files = []) {
    let li = '';

    for (let i = 0, len = files.length; i < len; i++) {
      let file = files[i];
      li += `<li>${file.name}（${bytesToSize(file.size)}）</li>`;
    }
    $(".files").html(li);
  }

  function handleFileInput() {
    console.log(g_xlsxData);
    // List column from input file:
    // Col 1: チャプター► サマリーPALMSレポート: "姓"    -- mapping --
    // Col 2: __EMPTY: "名"    -- mapping --
    // Col 3: __EMPTY_1: "出"    -- mapping --
    // Col 4: __EMPTY_2: "欠"    -- mapping --
    // Col 5: __EMPTY_4: "遅"    -- mapping --
    // Col 6: __EMPTY_5: "医"    -- mapping --
    // Col 7: __EMPTY_6: "代"    -- mapping --
    // Col 8: __EMPTY_8: "内与"    -- mapping --
    // Col 9: __EMPTY_9: "外与"    -- mapping --
    // Col 10: __EMPTY_10: "内受"    -- mapping --
    // Col 11: __EMPTY_12: "外受"    -- mapping --
    // Col 12: __EMPTY_13: "ビ"    -- mapping --
    // Col 13: __EMPTY_14: "1to1"    -- mapping --
    // Col 14: __EMPTY_15: "千円（千円未満切り捨て）"    -- mapping --
    // Col 15: __EMPTY_16: "CEU"    -- mapping --
    // Col 16: __EMPTY_17: "推薦状"    -- mapping --

    let fullName;
    let referral_final;
    let lastName; // Key: チャプター► サマリーPALMSレポート
    let firstName; // Key: __EMPTY
    let attendance; // Key: __EMPTY_1
    let absent; // Key: __EMPTY_2
    let late; // Key: __EMPTY_4
    let referral; // Key: __EMPTY_8
    let referral2; // Key: __EMPTY_9
    let visitor; // Key: __EMPTY_13
    let tyfcb; // Key: __EMPTY_15
    let ceu; // Key: __EMPTY_16
    let recommendation; // Key: __EMPTY_17
    let reportDate;

    let dataByDate = {};

    // Combine data from all files
    let data = [];
    for (let i = 0, len = g_xlsxData.length; i < len; i++) {
      let fileData = g_xlsxData[i];
      let _reportDate; // Key: __EMPTY_9

      // Get report date
      _reportDate = fileData[4];
      _reportDate = _reportDate[KEY.REPORT_DATE];
      
      if (!g_monthInReports.includes(_reportDate)) {
        g_monthInReports.push(_reportDate);
      }

      if (!dataByDate[_reportDate]) {
        dataByDate[_reportDate] = fileData;
      }

      // We don't need to calculate 3 item: 'ビジター', 'BNI', '合計'
      for (let j = 0, len2 = fileData.length - 3; j < len2; j++) {
        if (j < 7) continue;

        fileData[j][KEY.REPORT_DATE_NICE] = _reportDate;

        data.push(fileData[j]);
      }
    }

    console.log(data);

    // Calculate data of each member
    for (let i = 0, len = data.length; i < len; i++) {
      let row = data[i];

      lastName = row[Object.keys(row)[0]];
      firstName = row[KEY.FIRST_NAME];
      fullName = `${lastName} ${firstName}`;
      attendance = Number(row[KEY.ATTENDANCE]);
      absent = Number(row[KEY.ABSENT]);
      late = Number(row[KEY.LATE]);
      referral = Number(row[KEY.REFERRAL]);
      referral2 = Number(row[KEY.REFERRAL2]);
      referral_final = referral + referral2;
      visitor = Number(row[KEY.VISITOR]);
      tyfcb = Number(row[KEY.TYFCB]);
      ceu = Number(row[KEY.CEU]);
      recommendation = Number(row[KEY.RECOMMENDATION]);
      reportDate = row[KEY.REPORT_DATE_NICE];

      if (!g_members[fullName]) {
        g_members[fullName] = {};
      } else {
        attendance += g_members[fullName][KEY.ATTENDANCE_NICE];
        absent += g_members[fullName][KEY.ABSENT_NICE];
        late += g_members[fullName][KEY.LATE_NICE];
        referral += g_members[fullName][KEY.REFERRAL_NICE];
        referral2 += g_members[fullName][KEY.REFERRAL2_NICE];
        referral_final += g_members[fullName][KEY.REFERRAL_FINAL];
        visitor += g_members[fullName][KEY.VISITOR_NICE];
        tyfcb += g_members[fullName][KEY.TYFCB_NICE];
        ceu += g_members[fullName][KEY.CEU_NICE];
        recommendation += g_members[fullName][KEY.RECOMMENDATION_NICE];
      }

      g_members[fullName][KEY.LAST_NAME_NICE] = lastName;
      g_members[fullName][KEY.FIRST_NAME_NICE] = firstName;
      g_members[fullName][KEY.ATTENDANCE_NICE] = attendance;
      g_members[fullName][KEY.ABSENT_NICE] = absent;
      g_members[fullName][KEY.LATE_NICE] = late;
      g_members[fullName][KEY.REFERRAL_NICE] = referral;
      g_members[fullName][KEY.REFERRAL2_NICE] = referral2;
      g_members[fullName][KEY.REFERRAL_FINAL] = referral_final;
      g_members[fullName][KEY.VISITOR_NICE] = visitor;
      g_members[fullName][KEY.TYFCB_NICE] = tyfcb;
      g_members[fullName][KEY.CEU_NICE] = ceu;
      g_members[fullName][KEY.RECOMMENDATION_NICE] = recommendation;

      if (!g_members[fullName][KEY.REPORT_DATE_NICE]) {
        g_members[fullName][KEY.REPORT_DATE_NICE] = {};
      }

      if (!g_members[fullName][KEY.REPORT_DATE_NICE][reportDate]){
        g_members[fullName][KEY.REPORT_DATE_NICE][reportDate] = {};
      }

      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.ATTENDANCE_NICE] = Number(row[KEY.ATTENDANCE]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.ABSENT_NICE] = Number(row[KEY.ABSENT]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.LATE_NICE] = Number(row[KEY.LATE]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL_NICE] = Number(row[KEY.REFERRAL]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL2_NICE] = Number(row[KEY.REFERRAL2]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.REFERRAL_FINAL] = Number(row[KEY.REFERRAL]) + Number(row[KEY.REFERRAL2]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.VISITOR_NICE] = Number(row[KEY.VISITOR]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.TYFCB_NICE] = Number(row[KEY.TYFCB]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.CEU_NICE] = Number(row[KEY.CEU]);
      g_members[fullName][KEY.REPORT_DATE_NICE][reportDate][KEY.RECOMMENDATION_NICE] = Number(row[KEY.RECOMMENDATION]);
    }

    // Calculate and add rank score for each member
    const ruleOfScore = new RuleOfScore(KEY);

    for (const _fullName in g_members) {
      const member = g_members[_fullName];

      let earliestMonth = member[KEY.REPORT_DATE_NICE];

      const sorted = Object.keys(earliestMonth)
        .sort()
        .reduce((accumulator, key) => {
          accumulator[key] = earliestMonth[key];

          return accumulator;
        }, {});

      let firstMonth = sorted[Object.keys(sorted)[0]];

      // Using value in column ※1 to calculate rank score
      let absentTotal = member[KEY.ABSENT_NICE];
      let absentValRank = absentTotal - firstMonth[KEY.ABSENT_NICE];
      let lateTotal = member[KEY.LATE_NICE];
      let lateValRank = lateTotal - firstMonth[KEY.LATE_NICE];
      let referralTotal = member[KEY.REFERRAL_FINAL];
      let referralValRank = referralTotal - firstMonth[KEY.REFERRAL_FINAL];
      let visitorTotal = member[KEY.VISITOR_NICE];
      let visitorValRank = visitorTotal - firstMonth[KEY.VISITOR_NICE];
      let tyfcbTotal = member[KEY.TYFCB_NICE];
      let tyfcbValRank = tyfcbTotal - firstMonth[KEY.TYFCB_NICE];
      let ceuTotal = member[KEY.CEU_NICE];
      let ceuValRank = ceuTotal - firstMonth[KEY.CEU_NICE];
      let recommendationTotal = member[KEY.RECOMMENDATION_NICE];
      let recommendationValRank = recommendationTotal - firstMonth[KEY.RECOMMENDATION_NICE];

      // Rank of score table
      let absentScore = ruleOfScore.getScore(KEY.ABSENT, absentValRank);
      let lateScore = ruleOfScore.getScore(KEY.LATE, lateValRank);
      let referralScore = ruleOfScore.getScore(KEY.REFERRAL_FINAL, referralValRank / member[KEY.ATTENDANCE_NICE]);
      let visitorScore = ruleOfScore.getScore(KEY.VISITOR, visitorValRank);
      let tyfcbScore = ruleOfScore.getScore(KEY.TYFCB, tyfcbValRank);
      let ceuScore = ruleOfScore.getScore(KEY.CEU, ceuValRank);
      let recommendationScore = ruleOfScore.getScore(KEY.RECOMMENDATION, recommendationValRank);
      let totalScore = 0;
      totalScore = absentScore + lateScore + referralScore + visitorScore + tyfcbScore + ceuScore + recommendationScore;

      // Update current member with scoring values
      g_members[_fullName][KEY.ABSENT_SCORE] = absentScore;
      g_members[_fullName][KEY.LATE_SCORE] = lateScore;
      g_members[_fullName][KEY.REFERRAL_SCORE] = referralScore;
      g_members[_fullName][KEY.VISITOR_SCORE] = visitorScore;
      g_members[_fullName][KEY.TYFCB_SCORE] = tyfcbScore;
      g_members[_fullName][KEY.CEU_SCORE] = ceuScore;
      g_members[_fullName][KEY.RECOMMENDATION_SCORE] = recommendationScore;
      g_members[_fullName][KEY.TOTAL_SCORE] = totalScore;
    }

    g_members = sortObjects(g_members, KEY.TOTAL_SCORE, true, true);

    console.log(g_members);
  }

  function updateExportButtonState(isValid) {
    isValid = isValid || false;
    if (isValid) {
      $(".export-score").removeAttr("disabled");
    } else {
      $(".export-score").attr("disabled");
    }
  }

  function showMessage(msg) {
    if (!msg || msg == '') {
      $(".message").hide();
      return;
    }

    $(".message__text").text(msg);
    $(".message").show();
  }

  function readFile(file) {
    const xl2json = new ExcelToJSON();
    xl2json.parseExcel(file);
  }
  
  /**
   * Reset all input data when no selection upload action
   */
  function reset() {
    g_xlsxData = [];
    g_members = [];
    g_monthInReports = [];
    $(".files").html();
    $(".message").hide();
  }

  function exportXLSX() {
    // Don't export if empty data

    const ruleOfScore = new RuleOfScore(KEY);
    const worksheetData = [
      [
        '', // A
        '※1　今月なにも活動しない場合の合計数値', // B
      ], // Row 1
      [
        '', // A
        '', // B
        '', // C
        '参加', // D
        '欠席', '', // E F
        '遅刻', '', // G H
        'リファーラル', '', // I J
        'ビジター', '', // K L
        'TYFCB', '', // M N
        'CEU', '', // O P
        '推薦のことば', '', // Q R
        '', // S
        '今月何もしなかった場合の点数', // T
      ], // Row 2
      [
        '',
        '', 
        '', 
        '回数', // attendance
        '現在', '※1', // absent
        '現在', '※1', // late
        '現在', '※1', // referral
        '現在', '※1', // visitor
        '現在', '※1', // tyfcb
        '現在', '※1', // ceu
        '現在', '※1', // recommendation
        '', 
      ], // Row 3
      [
        '', 
        '姓', 
        '名', 
        '', // attendance
        '計', '', // absent
        '計', '', // late
        '計', '', // referral
        '計', '', // visitor
        '計', '', // tyfcb
        '計', '', // ceu
        '計', '', // recommendation

        '', // S

        '欠席', // T
        '遅刻', // U
        'リファーラル', // V
        'ビジター', // W
        'サンキュー金額', // X
        'CEU', // Y
        '推薦の言葉', // Z
        '点数', // AA
      ],
    ];

    for (const fullName in g_members) {
      const member = g_members[fullName];
      let row = [];

      let earliestMonth = member[KEY.REPORT_DATE_NICE];

      const sorted = Object.keys(earliestMonth)
        .sort()
        .reduce((accumulator, key) => {
          accumulator[key] = earliestMonth[key];

          return accumulator;
        }, {});

      let firstMonth = sorted[Object.keys(sorted)[0]];

      // Using value in column ※1 to calculate rank score
      let absentTotal = member[KEY.ABSENT_NICE];
      let absentValRank = absentTotal - firstMonth[KEY.ABSENT_NICE];
      let lateTotal = member[KEY.LATE_NICE];
      let lateValRank = lateTotal - firstMonth[KEY.LATE_NICE];
      let referralTotal = member[KEY.REFERRAL_FINAL];
      let referralValRank = referralTotal - firstMonth[KEY.REFERRAL_FINAL];
      let visitorTotal = member[KEY.VISITOR_NICE];
      let visitorValRank = visitorTotal - firstMonth[KEY.VISITOR_NICE];
      let tyfcbTotal = member[KEY.TYFCB_NICE];
      let tyfcbValRank = tyfcbTotal - firstMonth[KEY.TYFCB_NICE];
      let ceuTotal = member[KEY.CEU_NICE];
      let ceuValRank = ceuTotal - firstMonth[KEY.CEU_NICE];
      let recommendationTotal = member[KEY.RECOMMENDATION_NICE];
      let recommendationValRank = recommendationTotal - firstMonth[KEY.RECOMMENDATION_NICE];

      row.push(""); // A
      row.push(member[KEY.LAST_NAME_NICE]); // B
      row.push(member[KEY.FIRST_NAME_NICE]); // C

      row.push(member[KEY.ATTENDANCE_NICE]); // D

      // absent
      row.push(absentTotal);
      row.push(absentValRank); // Column ※1

      // late
      row.push(lateTotal);
      row.push(lateValRank); // Column ※1

      // referral
      row.push(referralTotal);
      row.push(referralValRank); // Column ※1

      // visitor
      row.push(visitorTotal);
      row.push(visitorValRank); // Column ※1

      // tyfcb
      row.push(tyfcbTotal);
      row.push(tyfcbValRank); // Column ※1

      // CEU
      row.push(ceuTotal);
      row.push(ceuValRank); // Column ※1

      // recommendation
      row.push(recommendationTotal);
      row.push(recommendationValRank); // Column ※1

      // Separate column
      row.push("");

      // Rank of score table
      let absentScore = ruleOfScore.getScore(KEY.ABSENT, absentValRank);
      let lateScore = ruleOfScore.getScore(KEY.LATE, lateValRank);
      let referralScore = ruleOfScore.getScore(KEY.REFERRAL_FINAL, referralValRank / member[KEY.ATTENDANCE_NICE]);
      let visitorScore = ruleOfScore.getScore(KEY.VISITOR, visitorValRank);
      let tyfcbScore = ruleOfScore.getScore(KEY.TYFCB, tyfcbValRank);
      let ceuScore = ruleOfScore.getScore(KEY.CEU, ceuValRank);
      let recommendationScore = ruleOfScore.getScore(KEY.RECOMMENDATION, recommendationValRank);

      row.push(absentScore);
      row.push(lateScore);
      row.push(referralScore);
      row.push(visitorScore);
      row.push(tyfcbScore);
      row.push(ceuScore);
      row.push(recommendationScore);

      let totalScore = 0;
      totalScore = absentScore + lateScore + referralScore + visitorScore + tyfcbScore + ceuScore + recommendationScore;

      row.push(totalScore);

      // Add row data into sheet
      worksheetData.push(row);
    }

    // Generate worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

    // Add merge
    let mergeAbsent = { s: { r:1, c:4 }, e: { r:1, c:5 } }; 
    // let mergeAbsent = XLSX.utils.decode_range("E2:F2"); // this is equivalent
    let mergeLate = XLSX.utils.decode_range("G2:H2");
    let mergeReferral = XLSX.utils.decode_range("I2:J2");
    let mergeVisitor = XLSX.utils.decode_range("K2:L2");
    let mergeTYFCB = XLSX.utils.decode_range("M2:N2");
    let mergeCEU = XLSX.utils.decode_range("O2:P2");
    let mergeRecommendation = XLSX.utils.decode_range("Q2:R2");
    let rankOfScoreTitle = XLSX.utils.decode_range("T2:AA3");

    if(!worksheet['!merges']) {
      worksheet['!merges'] = [];
    }

    worksheet['!merges'].push(mergeAbsent);
    worksheet['!merges'].push(mergeLate);
    worksheet['!merges'].push(mergeReferral);
    worksheet['!merges'].push(mergeVisitor);
    worksheet['!merges'].push(mergeTYFCB);
    worksheet['!merges'].push(mergeCEU);
    worksheet['!merges'].push(mergeRecommendation);
    worksheet['!merges'].push(rankOfScoreTitle);

    // Hint: If your worksheet data is auto generated and you don't know how many rows and columns are get populated 
    // then you could use the following way to find the number of rows and columns in the worksheet for doing cell width/height formatting.
    var range = XLSX.utils.decode_range(worksheet['!ref']);
    var noRows = range.e.r; // No.of rows
    var noCols = range.e.c; // No. of cols
    const textCenter = { alignment: { vertical: "center", horizontal: "center" } };

    for (let R = range.s.r; R <= noRows; ++R) {
      for (let C = range.s.c; C <= noCols; ++C) {
        let cell_address = { c:C, r:R };
        let cell_ref = XLSX.utils.encode_cell(cell_address);

        if (!worksheet[cell_ref]) continue;

        worksheet[cell_ref].s = {...textCenter };

        if (R > 3 && C == noCols) {
          let color = ruleOfScore.getRankColor( worksheet[cell_ref].v);
          worksheet[cell_ref].s.fill = { fgColor: { rgb: color }, };
        }

         // Background for column ※1
        if (R > 3 && (C == 5 || C == 7 || C == 9 || C == 11 || C == 13 || C == 15 || C == 17)) {
          worksheet[cell_ref].s.fill = { fgColor: { rgb: '8EAADB' }, };
        }

        // Set border for cell
        if ((R > 2 && ((C > 0 && C < 18) || (C > 18 && C <= noCols)))
          || (R > 1 && C > 2 && C < 18)
        ) {
          worksheet[cell_ref].s.border = { 
            top: { style: 'thin', color: '000000' },
            bottom: { style: 'thin', color: '000000' },
            left: { style: 'thin', color: '000000' },
            right: { style: 'thin', color: '000000' },
          };
        }

        // Set background color and font style for header
        if ((R == 3 && C > 0 && C < 18) || (R == 2 && C > 2 && C < 18)) {
          worksheet[cell_ref].s.fill = { fgColor: { rgb: 'BFBFBF' }, };
        }

        // Set font bold for header
        if (R == 3 && C > 0 && C < 18) {
          worksheet[cell_ref].s.font = { bold: true, sz: '14', };
        }
      }
    }

    // Set border for merged cells
    let bTopLeft = ['T2', 'D2', 'E2', 'G2', 'I2', 'K2', 'M2', 'O2', 'Q2'];
    let bTop = ['U2', 'V2', 'W2', 'X2', 'Y2', 'Z2'];
    let bTopRight = ['AA2', 'F2', 'H2', 'J2', 'L2', 'N2', 'P2', 'R2'];
    let bBottomRight = ['AA3'];
    let bBottom = ['U3', 'V3', 'W3', 'X3', 'Y3', 'Z3'];
    let bBottomLeft = ['T3'];
    
    bTopLeft.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { 
        border: { 
          top: { style: 'medium', color: '000000' },
          left: { style: 'medium', color: '000000' },
        }, 
        alignment: { vertical: "center", horizontal: "center" },
      };
    });
    bTop.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { border: { top: { style: 'medium', color: '000000' },}, };
    });
    bTopRight.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { 
        border: { 
          top: { style: 'medium', color: '000000' },
          right: { style: 'medium', color: '000000' },
        }, 
        alignment: { vertical: "center", horizontal: "center" },
      };
    });
    bBottomRight.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { 
        border: { 
          bottom: { style: 'medium', color: '000000' },
          right: { style: 'medium', color: '000000' },
        }, 
      };
    });
    bBottom.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { border: { bottom: { style: 'medium', color: '000000' },}, };
    });
    bBottomLeft.forEach(function(el) {
      if (!worksheet[el]) worksheet[el] = { s: {}, v: '', t: 's' };
      worksheet[el].s = { 
        border: { 
          bottom: { style: 'medium', color: '000000' },
          left: { style: 'medium', color: '000000' },
        }, 
      };
    });

    // Set color for rank title and align center
    worksheet['T2'].s.font = { color: { rgb: "FF0000" }, };
    worksheet['T2'].s.alignment = { vertical: "center", horizontal: "center" };

    // Set align left for B1
    worksheet['B1'].s.alignment = { vertical: "center", horizontal: "left" };
    worksheet['B1'].s.font = { sz: '14', };

    // Set width for columns
    var wscols = []
    for (let i = 0; i <= noCols; i++) {
      if (i == 0) {
        wscols.push({ wch: 2, });
      } else if (i == 1 || i == 2) {
        wscols.push({ wch: 10, });
      } else {
        wscols.push({ wch: 5, });
      }
    }

    worksheet['!cols'] = wscols;

    // Generate workbook
    const workbook = XLSX.utils.book_new();
    let sheetName = "Report";
    workbook.SheetNames.push(sheetName);
    workbook.Sheets[sheetName] = worksheet;

    // Write to xlsx file with file name
    XLSX.writeFile(workbook, `${getFirstDayOfCurrentMonth()}.xlsx`);
  }

  // TODO: do this later
  function generateWorkSheetData() {
    let wsData = [];

    g_members_sorted.forEach(function(member) {

    });

    return wsData;
  }

  class RuleOfScore {
    constructor(key) {
      this.KEY = key;
    }

    getScore = function(type, count) {
      switch (type) {
        case KEY.ABSENT:
          return this.getAbsentScore(count);
        case KEY.LATE:
          return this.getLateScore(count);
        case KEY.REFERRAL_FINAL:
          return this.getReferralScore(count);
        case KEY.VISITOR:
          return this.getVisitorScore(count);
        case KEY.TYFCB:
          return this.getTYFCBScore(count);
        case KEY.CEU:
          return this.getCEUScore(count);
        case KEY.RECOMMENDATION:
          return this.getRecommendationScore(count);
        default:
          break;
      }
    }

    getAbsentScore = function(count) {
      if (count >= 3) return 0;
      if (count == 2) return 5;
      if (count == 1) return 10;
      if (count == 0) return 15;
    }

    getLateScore = function(count) {
      if (count >= 1) return 0;
      if (count == 0) return 5;
    }

    getReferralScore = function(count) {
      if (count < 0.75) return 0;
      if (count >= 0.75 && count < 1) return 5;
      if (count >= 1 && count < 1.5) return 10;
      if (count >= 1.5 && count < 2) return 15;
      if (count >= 2) return 20;
    }

    getVisitorScore = function (count) {
      if (count < 2) return 0;
      if (count >= 2 && count < 4) return 5;
      if (count >= 4 && count < 6) return 10;
      if (count >= 6 && count < 12) return 15;
      if (count >= 12) return 20;
    }

    getTYFCBScore = function(count) {
      if (count < 200) return 0;
      if (count >= 200 && count < 500) return 5;
      if (count >= 500 && count < 1000) return 10;
      if (count >= 1000) return 15;
    }

    getCEUScore = function(count) {
      if (count < 10) return 0;
      if (count >= 10 && count < 15) return 5;
      if (count >= 15 && count < 20) return 10;
      if (count >= 20) return 15;
    }

    getRecommendationScore = function(count) {
      if (count == 0) return 0;
      if (count == 1) return 5;
      if (count >= 2) return 10;
    }

    getRankColor = function (score) {
      if (score == 100) {
        return "00B050"
      } else if (score >= 70 && score <= 95) {
        return "92D050";
      } else if (score >= 50 && score <= 65) {
        return "FFFF00";
      } else if (score >= 30 && score <= 45) {
        return "FF0000";
      } else if (score >= 0 && score <= 25) {
        return "BFBFBF";
      } else {
        // Do nothing
      }
    }
  }

  class ExcelToJSON {
    parseExcel = function (file) {
      var reader = new FileReader();

      reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {
          type: 'binary'
        });
        workbook.SheetNames.forEach(function (sheetName) {
          // Here is your object
          var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
          var json_object = JSON.stringify(XL_row_object);

          g_xlsxData.push(JSON.parse(json_object))
        });
      };

      reader.onerror = function (ex) {
        console.log(ex);
      };

      reader.readAsBinaryString(file);
    };
  }
})

/**
 * Get month from date string
 * @param   {string}  dateStr  Date string with format YYYY年MM月DD日
 * @return  {string}           return month with a character of month
 */
function extractMonth(dateStr) {
  const regex = /年\d+月/gm;

  // Alternative syntax using RegExp constructor
  // const regex = new RegExp('年\d+月', 'gm')

  let m;
  let month;

  while ((m = regex.exec(dateStr)) !== null) {
    // This is necessary to avoid infinite loops with zero-width matches
    if (m.index === regex.lastIndex) {
        regex.lastIndex++;
    }
    
    // The result can be accessed through the `m`-variable.
    m.forEach((match, groupIndex) => {
        month = match;
    });
  }

  month = month.replace("年", "");

  return month;
}

/**
 * Convert bytes to other unit KB, MB, GB, TB
 * @param  {number} bytes
 */
function bytesToSize(bytes) {
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
  if (bytes == 0) return '0 Byte';
  var i = parseInt(Math.floor(Math.log(bytes) / Math.log(1024)));
  return Math.round(bytes / Math.pow(1024, i), 2) + ' ' + sizes[i];
}

/**
 * Get the first day of month with format MMDD
 */
function getFirstDayOfCurrentMonth() {
  const d = new Date();
  let year = d.getFullYear();
  let month = d.getMonth() + 1;
  month = month.toString().padStart(2, '0');
  return `${year}${month}01`;
}

class Logger {
  constructor(debug) {
    this.debug = debug;
  }

  info = function(...data) {
    if (!this.debug) return;
    console.log(...data)
  }

  warn = function(...data) {
    if (!this.debug) return;
    console.warn(...data)
  }
}

/**
 * Sort object properties (only own properties will be sorted).
 * @param {object} obj object to sort properties
 * @param {string|int} sortedBy 1 - sort object properties by specific value.
 * @param {bool} isNumericSort true - sort object properties as numeric value, false - sort as string value.
 * @param {bool} reverse false - reverse sorting.
 * @returns {Array} array of items in [[key,value],[key,value],...] format.
 */
function sortProperties(obj, sortedBy, isNumericSort, reverse) {
  sortedBy = sortedBy || 1; // by default first key
  isNumericSort = isNumericSort || false; // by default text sort
  reverse = reverse || false; // by default no reverse

  var reversed = (reverse) ? -1 : 1;

  var sortable = [];
  for (var key in obj) {
      if (obj.hasOwnProperty(key)) {
          sortable.push([key, obj[key]]);
      }
  }
  if (isNumericSort)
      sortable.sort(function (a, b) {
          return reversed * (a[1][sortedBy] - b[1][sortedBy]);
      });
  else
      sortable.sort(function (a, b) {
          var x = a[1][sortedBy].toLowerCase(),
              y = b[1][sortedBy].toLowerCase();
          return x < y ? reversed * -1 : x > y ? reversed : 0;
      });
  return sortable; // array in format [ [ key1, val1 ], [ key2, val2 ], ... ]
}


function sortObjects(objects, sortedBy, isNumericSort, reverse) {
  var newObject = {};
  var sortedArray = sortProperties(objects, sortedBy, isNumericSort, reverse);

  for (const element of sortedArray) {
      var key = element[0];
      var value = element[1];

      newObject[key] = value;
  }

  return newObject;
}
