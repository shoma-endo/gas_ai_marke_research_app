// 設定オブジェクト
const CONFIG = {
  apiKey: PropertiesService.getScriptProperties().getProperty("apiKey"),
  apiUrl: "https://api.openai.com/v1/chat/completions",
  sheet: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(),
  headers: {
    Authorization: "",
    "Content-type": "application/json",
  },
};

// APIキーの設定
CONFIG.headers.Authorization = `Bearer ${CONFIG.apiKey}`;

// シートからデータを取得
function getSheetData() {
  return {
    brandName: CONFIG.sheet.getRange("A2").getValue(),
    serviceExplanation: CONFIG.sheet.getRange("B2").getValue(),
  };
}

// 先頭のハイフンを削除
function trimLeadingHyphen(str) {
  return str.startsWith("- ") ? str.substring(2) : str;
}

// OpenAIからコンテンツを取得
function fetchOpenAIContent(promptContent) {
  const systemMessage = `あなたは、ゴールドマン・サックス、マッキンゼー、モルガン・ス タンレー、ベイン、 PWC、 BCG、 P&G、 アクセンチュアが合併した コンサル会社のパートナー・コンサルタントです。
    パートナー・コンサルタントとして、 必要に応じて適切なフレーム ワークやテクニックを用いながら、ユーザーの質問に解答する際にはデータのみを返却してください。
    スプレッドシートのセルとして使用するため、顧客セグメントの情報を1つだけ提供してください。
    それ以外の余計な情報は提供しないでください。
    あなたはユーザーの学び、利益、出世など、ユーザー便益の最大化を目指す応対を行います。`;

  const messages = [
    { role: "system", content: systemMessage },
    { role: "user", content: promptContent },
  ];

  const options = {
    muteHttpExceptions: true,
    headers: CONFIG.headers,
    method: "POST",
    payload: JSON.stringify({
      model: "gpt-3.5-turbo",
      temperature: 0.9,
      messages: messages,
    }),
  };

  const response = JSON.parse(
    UrlFetchApp.fetch(CONFIG.apiUrl, options).getContentText()
  );
  // trimLeadingHyphen()を使用して、コンテンツの先頭から"- "を削除
  return trimLeadingHyphen(response.choices[0].message.content);
}

// 列の設定を行う
function setupColumns() {
  // 列C〜Kのヘッダーに値を設定
  CONFIG.sheet.getRange("C1").setValue("顧客セグメント");
  CONFIG.sheet.getRange("D1").setValue("ペルソナ");
  CONFIG.sheet.getRange("E1").setValue("インサイト");
  CONFIG.sheet.getRange("F1").setValue("便益");
  CONFIG.sheet.getRange("G1").setValue("コピー＿インサイト考慮なし");
  CONFIG.sheet.getRange("H1").setValue("コピー＿インサイト考慮あり");
  CONFIG.sheet.getRange("I1").setValue("企画案");
  CONFIG.sheet.getRange("J1").setValue("施策案＿Paid・広告");
  CONFIG.sheet.getRange("K1").setValue("施策案＿Owend");

  // 列C〜Eの背景色を設定
  CONFIG.sheet.getRange("C1:E1").setBackground("#FEEFC9");

  // 列F〜Hの背景色を設定
  CONFIG.sheet.getRange("F1:H1").setBackground("#D3E6CE");

  // 列I〜Kの背景色を設定 (ソフトラベンダーカラーを選択)
  CONFIG.sheet.getRange("I1:K1").setBackground("#D9D2E9");

  // 列C〜Kの幅を設定
  CONFIG.sheet.setColumnWidth(3, 300); // 3 represents column C
  CONFIG.sheet.setColumnWidth(4, 300); // 4 represents column D
  CONFIG.sheet.setColumnWidth(5, 300); // 5 represents column E
  CONFIG.sheet.setColumnWidth(6, 300); // 6 represents column F
  CONFIG.sheet.setColumnWidth(7, 300); // 7 represents column G
  CONFIG.sheet.setColumnWidth(8, 300); // 8 represents column H
  CONFIG.sheet.setColumnWidth(9, 300); // 9 represents column I
  CONFIG.sheet.setColumnWidth(10, 300); // 10 represents column J
  CONFIG.sheet.setColumnWidth(11, 300); // 11 represents column K

  // 列C〜Kのヘッダーのフォントを太字に設定
  CONFIG.sheet.getRange("C1:K1").setFontWeight("bold");

  // テキストの折り返しを設定
  CONFIG.sheet.getRange("A1:Z50").setWrap(true);
}

function generateCustomerSegmentData(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();

  const prompt = `
  ### 依頼: ###
  あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスを売るマーケティング戦略設計コンサルタントです。現在、マーケティング戦略の一環として、ターゲット顧客セグメントの具体的なグループや特性を特定する必要があります。このタスクのために、下記のフォーマットで「顧客セグメント」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。

  期待する出力フォーマット: ###
  グループ名：グループの特性や特徴を含む簡単な説明
  ###

  例: ###
  - 若者：新しいエンターテイメントを探求する20代学生や新社会人
  - 都心勤務者：都心で働くビジネスパーソンや専門職の社会人
  ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Customer Segment for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 3).setValue(content); // Assumes column 3 (C) is where "顧客セグメント" data is stored
}

function generatePersonaData(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C

  const prompt = `
    ### 依頼: ###
    あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。提供された「顧客セグメント」(${customerSegment})を元に、そのセグメントに対応する「ペルソナ」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
  
    期待する出力フォーマット: ###
    ペルソナ名：ペルソナのライフスタイル、興味、関心などの詳細な説明
    ###
  
    例: ###
    - 学生太郎：都心の大学に通う22歳の男性。週末は友人と映画を観たり、新しいカフェを探しに行ったりする。スマートフォンを頻繁に使い、新しいテクノロジーやアプリに興味がある。
    ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Persona for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 4).setValue(content); // Assumes column 4 (D) is where "ペルソナ" data is stored
}

function generateInsightData(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D

  const prompt = `
    ### 依頼: ###
    あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。提供された「顧客セグメント」(${customerSegment})と「ペルソナ」(${persona})を元に、そのターゲット顧客の隠れたニーズや心理、行動を反映した「インサイト」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
  
    期待する出力フォーマット: ###
    ターゲット顧客の隠れたニーズや心理、行動に関する具体的な情報
    ###
  
    例: ###
    - 新しいトレンドを追いたい、友人との会話のネタを探している、忙しい日常の中で気軽にエンターテイメントを楽しみたい
    ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Insight for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 5).setValue(content); // Assumes column 5 (E) is where "インサイト" data is stored
}

function generateBenefitData(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const insight = CONFIG.sheet.getRange(rowIndex, 5).getValue(); // Get value from column E

  const prompt = `
      ### 依頼: ###
      あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、及び「インサイト」(${insight})を元に、そのターゲット顧客にとっての「便益」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
    
      期待する出力フォーマット: ###
      ターゲット顧客にとってのサービスや商品の主な利点やメリット
      ###
    
      例: ###
      - 最新トレンドのエンターテイメントを取り入れ、常に新鮮な体験を提供。気軽にエンターテイメントを楽しむことができる。
      ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Benefit for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 6).setValue(content); // Assumes column 6 (F) is where "便益" data is stored
}

function generateCopyWithoutInsight(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const benefit = CONFIG.sheet.getRange(rowIndex, 6).getValue(); // Get value from column F

  const prompt = `
      ### 依頼: ###
      あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、及び「便益」(${benefit})を元に、インサイトを考慮せずに、そのサービスや商品の主なメッセージを伝える「コピー＿インサイト考慮なし」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
    
      期待する出力フォーマット: ###
      ターゲット顧客にとってのサービスや商品の主なメッセージ
      ###
    
      例: ###
      - 「新しい体験、ここで待っています」
      ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Copy without Insight for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 7).setValue(content); // Assumes column 7 (G) is where "コピー＿インサイト考慮なし" data is stored
}

function generateCopyWithInsight(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const insight = CONFIG.sheet.getRange(rowIndex, 5).getValue(); // Get value from column E
  const benefit = CONFIG.sheet.getRange(rowIndex, 6).getValue(); // Get value from column F

  const prompt = `
      ### 依頼: ###
      あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、「インサイト」(${insight})及び「便益」(${benefit})を元に、インサイトを考慮した、そのサービスや商品の主なメッセージを伝える「コピー＿インサイト考慮あり」の情報を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
    
      期待する出力フォーマット: ###
      ターゲット顧客にとってのサービスや商品の主なメッセージ、インサイトを反映した内容
      ###
    
      例: ###
      - 「あなたの求めている新しい体験、ここで実現します」
      ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Copy with Insight for row ${rowIndex}`, content);

  // Write the received content to the specific cell in the sheet
  CONFIG.sheet.getRange(rowIndex, 8).setValue(content); // Assumes column 8 (H) is where "コピー＿インサイト考慮あり" data is stored
}

function generateProposal(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const insight = CONFIG.sheet.getRange(rowIndex, 5).getValue(); // Get value from column E
  const benefit = CONFIG.sheet.getRange(rowIndex, 6).getValue(); // Get value from column F

  const prompt = `
        ### 依頼: ###
        あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。
        提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、「インサイト」(${insight})及び「便益」(${benefit})を元に、企画案を考えてください。
        100文字以内で、企画案の簡単な説明を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
      
        期待する出力フォーマット: ###
        企画案の簡単な説明
        ###
      
        例: ###
        - 20代の学生や新社会人をターゲットに、最新のエンターテイメントを提供するサービスを展開します。新しいトレンドを取り入れ、常に新鮮な体験を提供することで、ターゲットのニーズを満たします。
        ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Proposal for row ${rowIndex}`, content);

  // Assuming column 9 (I) is where "企画案" data is stored
  CONFIG.sheet.getRange(rowIndex, 9).setValue(content);
}

function generatePaidAdProposal(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const insight = CONFIG.sheet.getRange(rowIndex, 5).getValue(); // Get value from column E
  const benefit = CONFIG.sheet.getRange(rowIndex, 6).getValue(); // Get value from column F

  const prompt = `
          ### 依頼: ###
          あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。
          提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、「インサイト」(${insight})及び「便益」(${benefit})を元に、Paid・広告の企画案を考えてください。
          100文字以内で、Paid・広告の企画案の簡単な説明を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
        
          期待する出力フォーマット: ###
          Paid・広告の詳細な説明
          ###
        
          例: ###
          - ソーシャルメディア広告を利用して20代の学生や新社会人をターゲットに、最新のエンターテイメントを提供するサービスを展開します。新しいトレンドを取り入れ、常に新鮮な体験を提供することで、ターゲットのニーズを満たします。
          ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Paid Ad Proposal for row ${rowIndex}`, content);

  // Assuming column 10 (J) is where "施策案＿Paid・広告" data is stored
  CONFIG.sheet.getRange(rowIndex, 10).setValue(content);
}

function generateOwnedProposal(rowIndex) {
  const { brandName, serviceExplanation } = getSheetData();
  const customerSegment = CONFIG.sheet.getRange(rowIndex, 3).getValue(); // Get value from column C
  const persona = CONFIG.sheet.getRange(rowIndex, 4).getValue(); // Get value from column D
  const insight = CONFIG.sheet.getRange(rowIndex, 5).getValue(); // Get value from column E
  const benefit = CONFIG.sheet.getRange(rowIndex, 6).getValue(); // Get value from column F

  const prompt = `
        ### 依頼: ###
        あなたは、「${brandName}」というブランド名の「${serviceExplanation}」というサービスのマーケティング戦略設計コンサルタントです。
        提供された「顧客セグメント」(${customerSegment})、「ペルソナ」(${persona})、「インサイト」(${insight})及び「便益」(${benefit})を元に、Owendメディアの企画案を考えてください。
       
        100文字以内で、Owendメディアの企画案の簡単な説明を1つだけ提供してください。余計な説明やフレーズは追加しないでください。
        期待する出力フォーマット: ###
        Owendメディアの詳細な説明
        ###
        
        例: ###
        - ブランドの公式ウェブサイトやSNSアカウントを通じて、ターゲットに合わせたコンテンツを提供。月に一度のニュースレターを通じて新しい情報やプロモーションを伝える。
        ###`;

  console.log("prompt", prompt);

  const content = fetchOpenAIContent(prompt);
  console.log(`Owned Proposal for row ${rowIndex}`, content);

  // Assuming column 11 (K) is where "施策案＿Owend" data is stored
  CONFIG.sheet.getRange(rowIndex, 11).setValue(content);
}

function runSteps() {
  setupColumns();

  for (let i = 2; i <= 6; i++) {
    console.log(`Processing row ${i}`);
    generateCustomerSegmentData(i);
    generatePersonaData(i);
    generateInsightData(i);
    generateBenefitData(i);
    generateCopyWithoutInsight(i);
    generateCopyWithInsight(i);
    generateProposal(i);
    generatePaidAdProposal(i);
    generateOwnedProposal(i);
    console.log(`Finished processing row ${i}`);
  }
}

