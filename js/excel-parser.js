// excel-parser.js - Excel & CSV Parser with intelligent column mapping & multi-sheet support

class ExcelParser {
  static parseFile(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          
          const result = ExcelParser.processWorkbook(workbook, file.name);
          resolve(result);
        } catch (err) {
          reject(new Error("파일 파싱 중 오류가 발생했습니다: " + err.message));
        }
      };

      reader.onerror = () => reject(new Error("파일을 읽는 중 오류가 발생했습니다."));
      reader.readAsArrayBuffer(file);
    });
  }

  static processWorkbook(workbook, fileName) {
    const sheetNames = workbook.SheetNames;
    if (!sheetNames || sheetNames.length === 0) {
      throw new Error("엑셀 파일에 시트가 존재하지 않습니다.");
    }

    // Sheet 1 (Main projects list)
    const sheet1 = workbook.Sheets[sheetNames[0]];
    const rawSheet1 = XLSX.utils.sheet_to_json(sheet1, { header: 1, defval: "" });

    // Detect header row (first 10 rows)
    const headerInfo1 = ExcelParser.findHeaderRow(rawSheet1);
    if (!headerInfo1) {
      throw new Error("헤더 행을 찾을 수 없습니다. 컬럼 구성을 확인해주세요.");
    }

    const { headerRowIndex, headers } = headerInfo1;
    const dataRows1 = rawSheet1.slice(headerRowIndex + 1);

    // Parse Sheet 2 (Detail information if exists)
    let detailsMap = new Map();
    if (sheetNames.length > 1) {
      const sheet2 = workbook.Sheets[sheetNames[1]];
      const rawSheet2 = XLSX.utils.sheet_to_json(sheet2, { header: 1, defval: "" });
      const headerInfo2 = ExcelParser.findHeaderRow(rawSheet2);
      if (headerInfo2) {
        const detailRows = rawSheet2.slice(headerInfo2.headerRowIndex + 1);
        const map2 = ExcelParser.mapHeaders(headerInfo2.headers);
        
        detailRows.forEach(row => {
          if (!row || row.length === 0) return;
          const noVal = map2.No !== -1 ? String(row[map2.No] || "").trim() : "";
          const nameVal = map2.Name !== -1 ? String(row[map2.Name] || "").trim() : "";
          const key = noVal || nameVal;
          if (key) {
            detailsMap.set(key, {
              Target: map2.Target !== -1 ? String(row[map2.Target] || "").trim() : "",
              Basis: map2.Basis !== -1 ? String(row[map2.Basis] || "").trim() : "",
              Detail: map2.Detail !== -1 ? String(row[map2.Detail] || "").trim() : "",
              Effect: map2.Effect !== -1 ? String(row[map2.Effect] || "").trim() : ""
            });
          }
        });
      }
    }

    // Map columns for Sheet 1
    const colMap = ExcelParser.mapHeaders(headers);

    const projects = [];
    let autoIndex = 1;

    dataRows1.forEach((row) => {
      if (!row || row.length === 0) return;

      // Check if row has valid content (at least name or budget)
      const nameVal = colMap.Name !== -1 ? String(row[colMap.Name] || "").trim() : "";
      if (!nameVal && (!row[colMap.Y27] && !row[colMap.Dept])) return;

      const noVal = colMap.No !== -1 ? row[colMap.No] : autoIndex;
      const origNoVal = colMap.OrigNo !== -1 ? String(row[colMap.OrigNo] || "").trim() : "";
      const deptVal = colMap.Dept !== -1 ? String(row[colMap.Dept] || "").trim() : "";
      let categoryVal = colMap.Category !== -1 ? String(row[colMap.Category] || "").trim() : "";
      const contactVal = colMap.Contact !== -1 ? String(row[colMap.Contact] || "").trim() : "";
      const y26Val = colMap.Y26 !== -1 ? ExcelParser.cleanNumber(row[colMap.Y26]) : 0;
      const y27Val = colMap.Y27 !== -1 ? ExcelParser.cleanNumber(row[colMap.Y27]) : 0;
      let diffVal = colMap.Diff !== -1 ? ExcelParser.cleanNumber(row[colMap.Diff]) : (y27Val - y26Val);
      const unitVal = colMap.Unit !== -1 ? String(row[colMap.Unit] || "억원").trim() : "억원";
      const noteVal = colMap.Note !== -1 ? String(row[colMap.Note] || "").trim() : "";

      // Infer agency
      const agencyVal = ExcelParser.inferAgency(deptVal);

      // Infer category if missing
      if (!categoryVal) {
        categoryVal = ExcelParser.inferCategory(nameVal, deptVal, agencyVal);
      }

      // Check detail map
      const lookupKey = String(noVal).trim() || nameVal;
      const extraDetails = detailsMap.get(lookupKey) || detailsMap.get(nameVal) || {};

      const project = {
        No: isNaN(Number(noVal)) ? autoIndex : Number(noVal),
        OrigNo: origNoVal,
        Name: nameVal || `신규사업 #${autoIndex}`,
        Dept: deptVal,
        Agency: agencyVal,
        Category: categoryVal,
        Contact: contactVal,
        Y26: y26Val,
        Y27: y27Val,
        Diff: diffVal,
        Unit: unitVal,
        Note: noteVal,
        Target: extraDetails.Target || (colMap.Target !== -1 ? String(row[colMap.Target] || "").trim() : ""),
        Basis: extraDetails.Basis || (colMap.Basis !== -1 ? String(row[colMap.Basis] || "").trim() : ""),
        Detail: extraDetails.Detail || (colMap.Detail !== -1 ? String(row[colMap.Detail] || "").trim() : ""),
        Effect: extraDetails.Effect || (colMap.Effect !== -1 ? String(row[colMap.Effect] || "").trim() : "")
      };

      projects.push(project);
      autoIndex++;
    });

    const totalBudget = projects.reduce((acc, p) => acc + (p.Diff > 0 ? p.Diff : (p.Y27 || 0)), 0);

    return {
      fileName: fileName,
      sheetNames: sheetNames,
      headers: headers,
      colMap: colMap,
      projects: projects,
      summary: {
        title: fileName.replace(/\.[^/.]+$/, "") + " 현황",
        subtitle: `총 ${projects.length}개 사업 분석 결과`,
        total_count: projects.length,
        total_budget: Math.round(totalBudget * 100) / 100,
        unit: projects[0]?.Unit || "억원"
      }
    };
  }

  static findHeaderRow(rows) {
    if (!rows || rows.length === 0) return null;

    // Scan up to 10 rows
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const row = rows[i];
      if (!row || !Array.isArray(row)) continue;

      const rowStr = row.map(c => String(c || "").trim().toLowerCase()).join(" ");
      
      // Keywords that strongly indicate header row
      const keywords = ["연번", "사업명", "예산", "부서", "기관", "증감", "내용", "no", "name", "budget", "dept"];
      let matchCount = 0;
      keywords.forEach(kw => {
        if (rowStr.includes(kw)) matchCount++;
      });

      if (matchCount >= 2) {
        return {
          headerRowIndex: i,
          headers: row.map(c => String(c || "").trim())
        };
      }
    }

    // Default to row 0 if non-empty
    if (rows[0] && rows[0].length > 0) {
      return {
        headerRowIndex: 0,
        headers: rows[0].map(c => String(c || "").trim())
      };
    }

    return null;
  }

  static mapHeaders(headers) {
    const map = {
      No: -1,
      OrigNo: -1,
      Name: -1,
      Dept: -1,
      Category: -1,
      Contact: -1,
      Y26: -1,
      Y27: -1,
      Diff: -1,
      Unit: -1,
      Note: -1,
      Target: -1,
      Basis: -1,
      Detail: -1,
      Effect: -1
    };

    headers.forEach((h, idx) => {
      const col = h.replace(/\s+/g, "").toLowerCase();

      if (map.No === -1 && /^(연번|순번|번호|no|id)$/.test(col)) {
        map.No = idx;
      } else if (map.OrigNo === -1 && /원사업번호|원번호|사업번호|기존번호/.test(col)) {
        map.OrigNo = idx;
      } else if (map.Name === -1 && /(신규사업명|사업명|과제명|사업|name|project)/.test(col)) {
        map.Name = idx;
      } else if (map.Category === -1 && /(분야|중점투자분야|투자분야|분야별|부문|category|area)/.test(col)) {
        map.Category = idx;
      } else if (map.Dept === -1 && /(주관기관|담당부서|주관부서|소관부서|담당과|소관부처|기관|dept|department)/.test(col)) {
        map.Dept = idx;
      } else if (map.Contact === -1 && /(연락처|담당자|contact|phone|tel)/.test(col)) {
        map.Contact = idx;
      } else if (map.Y26 === -1 && /(26년|'26년|2026년|전년예산|당초예산|2026년예산)/.test(col)) {
        map.Y26 = idx;
      } else if (map.Y27 === -1 && /(27년정부안|'27년정부안|2027년정부안|27년예산|'27년예산|2027년예산|정부안\(신규\)|신규예산|정부안|예산|금액|budget)/.test(col)) {
        map.Y27 = idx;
      } else if (map.Diff === -1 && /(신규증감액|증감액|신규증감|증감|diff|increase)/.test(col)) {
        map.Diff = idx;
      } else if (map.Unit === -1 && /(단위|unit)/.test(col)) {
        map.Unit = idx;
      } else if (map.Note === -1 && /(핵심신규사업내용|핵심사업내용|추진목표|사업내용|세부내용|비고|적요|note|description)/.test(col)) {
        map.Note = idx;
      } else if (map.Target === -1 && /(지원대상|수혜규모|지원대상및수혜규모|대상)/.test(col)) {
        map.Target = idx;
      } else if (map.Basis === -1 && /(추진근거|관련법령|근거|법령)/.test(col)) {
        map.Basis = idx;
      } else if (map.Detail === -1 && /(세부추진내용|도입일정|세부내용및도입일정)/.test(col)) {
        map.Detail = idx;
      } else if (map.Effect === -1 && /(기대효과|정책의의|기대효과및정책의의|효과)/.test(col)) {
        map.Effect = idx;
      }
    });

    return map;
  }

  static cleanNumber(val) {
    if (val === null || val === undefined || val === "" || val === "-") return 0;
    if (typeof val === "number") return isNaN(val) ? 0 : val;
    
    // String cleanup
    const str = String(val).replace(/,/g, "").replace(/억\s*원?|원|명/g, "").trim();
    const num = parseFloat(str);
    return isNaN(num) ? 0 : num;
  }

  static inferAgency(dept) {
    if (!dept) return "기타";
    if (dept.includes("방위사업청") || dept.includes("방사청")) return "방위사업청";
    if (dept.includes("국방부")) return "국방부";
    if (dept.includes("병무청")) return "병무청";
    if (dept.includes("합참") || dept.includes("합동참모본부")) return "합참";
    if (dept.includes("육군")) return "육군";
    if (dept.includes("해군")) return "해군";
    if (dept.includes("공군")) return "공군";
    if (dept.includes("해병대")) return "해병대";

    // Split by space and take first token
    const firstWord = dept.split(/\s+/)[0];
    return firstWord || "기타";
  }

  static inferCategory(name, dept, agency) {
    const text = (name + " " + dept).toLowerCase();

    if (text.includes("반도체") || text.includes("항공엔진") || text.includes("r&d") || text.includes("연구개발") || text.includes("첨단무기") || text.includes("우주") || text.includes("미래기술")) {
      return "첨단무기 R&D";
    }
    if (text.includes("간부") || text.includes("기본급") || text.includes("적금") || text.includes("처우") || text.includes("봉급") || text.includes("숙소") || text.includes("주거") || text.includes("보험") || text.includes("복지")) {
      return "장병 복지·처우개선";
    }
    if (text.includes("예비군") || text.includes("동원") || text.includes("훈련") || text.includes("의료팀")) {
      return "예비전력 정예화 및 안전";
    }
    if (text.includes("ai") || text.includes("생성형") || text.includes("원스톱") || text.includes("시스템") || text.includes("정보화") || text.includes("디지털")) {
      return "스마트 행정·AI";
    }
    if (text.includes("아웃소싱") || text.includes("조직") || text.includes("제도") || text.includes("시범사업")) {
      return "군 조직·운영 혁신";
    }
    if (text.includes("병역") || text.includes("입영") || text.includes("항공여비") || text.includes("국외")) {
      return "병무 행정·복무지원";
    }

    if (agency === "방위사업청") return "방위사업·R&D";
    if (agency === "병무청") return "병무 행정";
    if (agency === "국방부") return "국방 정책·전력";

    return "일반 행정";
  }
}

window.ExcelParser = ExcelParser;
