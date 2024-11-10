import React, {useEffect, useState} from "react";
import * as XLSX from "xlsx";
import "./ExcelMergeApp.css";

function ExcelMergeApp() {
	const [onchData, setOnchData] = useState([]);
	const [coupangData, setCoupangData] = useState([]);
	const [mergedPreview, setMergedPreview] = useState([]);
	const [isDarkMode, setIsDarkMode] = useState(true);

	useEffect(() => {
		if (isDarkMode) {
			document.body.classList.add("dark-mode");
			document.documentElement.classList.add("dark-mode");
		} else {
			document.body.classList.remove("dark-mode");
			document.documentElement.classList.remove("dark-mode");
		}
	}, [isDarkMode]);

	const handleFileUpload = (e, setData, type) => {
		const file = e.target.files[0];
		const reader = new FileReader();
		reader.onload = (event) => {
			const data = new Uint8Array(event.target.result);
			const workbook = XLSX.read(data, {type: "array"});
			const sheetName = workbook.SheetNames[0];
			const sheet = workbook.Sheets[sheetName];
			const jsonData = XLSX.utils.sheet_to_json(sheet);

			const cleanedData = jsonData.map(row => {
				if (type === "onch") {
					const {온채널주문코드, 상품명, 상품코드, 옵션, 수량, 가격} = row;
					return {온채널주문코드, 온채널상품명: 상품명, 온채널상품코드: 상품코드, 옵션, 수량, 가격};
				} else if (type === "coupang") {
					const {주문번호, 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비} = row;
					return {쿠팡주문번호: 주문번호, 쿠팡상품번호: 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비};
				}
				return row;
			});

			setData(cleanedData);
		};
		reader.readAsArrayBuffer(file);
	};

	const mergeData = () => {
		const maxLength = Math.max(onchData.length, coupangData.length);
		const mergedData = [];

		for (let i = 0; i < maxLength; i++) {
			const {온채널주문코드, 온채널상품명: 상품명, 온채널상품코드: 상품코드, 옵션, 수량, 가격} = onchData[i] || {};
			const {쿠팡주문번호: 주문번호, 쿠팡상품번호: 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비} = coupangData[i] || {};

			const 수수료 =  Math.round(판매금액 * 0.108)
			const 판매 = 판매금액 + 판매배송비;
			const 순이익 = 판매 - 가격 - 수수료 * (수량 || 1);

			const mergedRow = {
				온채널주문코드,
				온채널상품명: 상품명,
				온채널상품코드: 상품코드,
				옵션,
				쿠팡주문번호: 주문번호,
				쿠팡상품번호: 상품번호,
				수량,
				도매: 가격,
				판매: 판매,
				취소금액,
				취소배송비,
				예상수수료: 수수료,
				순이익,
			};

			mergedData.push(mergedRow);
		}


		const summaryRow = {
			온채널주문코드: "총합계",
			온채널상품명: "-",
			온채널상품코드: "-",
			옵션: "-",
			쿠팡주문번호: "-",
			쿠팡상품번호: "-",
			수량: mergedData.reduce((sum, row) => sum + (row.수량 || 0), 0),
			도매: mergedData.reduce((sum, row) => sum + (row.도매 || 0), 0),
			판매: mergedData.reduce((sum, row) => sum + (row.판매 || 0), 0),
			취소금액: mergedData.reduce((sum, row) => sum + (row.취소금액 || 0), 0),
			취소배송비: mergedData.reduce((sum, row) => sum + (row.취소배송비 || 0), 0),
			예상수수료: mergedData.reduce((sum, row) => sum + (row.예상수수료 || 0), 0),
			순이익: mergedData.reduce((sum, row) => sum + (row.순이익 || 0), 0),
		};
		mergedData.push(summaryRow);

		setMergedPreview(mergedData);
	};

	const downloadMergedData = () => {
		mergeData(); // 병합 데이터를 미리보기와 동시에 다운로드에 사용

		const worksheet = XLSX.utils.json_to_sheet(mergedPreview);
		const workbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(workbook, worksheet, "MergedData");
		XLSX.writeFile(workbook, "MergedData.xlsx");
	};

	return (
		<div className={`app-container ${isDarkMode ? "dark-mode" : ""}`}>
			<h1>Excel Merge App</h1>
			<button className="toggle-mode" onClick={() => setIsDarkMode(!isDarkMode)}>
				{isDarkMode ? "Light Mode" : "Dark Mode"}
			</button>
			<div className="file-input">
				<p>온채널 엑셀 파일 업로드:</p>
				<input type="file" accept=".xlsx, .xls" onChange={(e) => handleFileUpload(e, setOnchData, "onch")}/>
			</div>
			<div className="file-input">
				<p>쿠팡 엑셀 파일 업로드:</p>
				<input type="file" accept=".xlsx, .xls" onChange={(e) => handleFileUpload(e, setCoupangData, "coupang")}/>
			</div>
			<button className="merge-btn" onClick={mergeData}>Preview Merged Data</button>
			<button className="download-btn" onClick={downloadMergedData}>Download Merged Excel</button>

			{mergedPreview.length > 0 && (
				<div className="preview-container">
					<h2>미리보기</h2>
					<table className="preview-table">
						<thead>
						<tr>
							{Object.keys(mergedPreview[0]).map((key) => (
								<th key={key}>{key}</th>
							))}
						</tr>
						</thead>
						<tbody>
						{mergedPreview.map((row, index) => (
							<tr key={index}>
								{Object.values(row).map((value, i) => (
									<td key={i}>{value}</td>
								))}
							</tr>
						))}
						</tbody>
					</table>
				</div>
			)}
		</div>
	);
}
export default ExcelMergeApp;
