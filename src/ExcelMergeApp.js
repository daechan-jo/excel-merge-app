import React, { useState } from "react";
import * as XLSX from "xlsx";

function ExcelMergeApp() {
	const [onchData, setOnchData] = useState([]);
	const [coupangData, setCoupangData] = useState([]);

	// 파일 업로드 후 데이터 파싱 및 정리
	const handleFileUpload = (e, setData, type) => {
		const file = e.target.files[0];
		const reader = new FileReader();
		reader.onload = (event) => {
			const data = new Uint8Array(event.target.result);
			const workbook = XLSX.read(data, { type: "array" });
			const sheetName = workbook.SheetNames[0];
			const sheet = workbook.Sheets[sheetName];
			const jsonData = XLSX.utils.sheet_to_json(sheet);

			// 파일 유형에 따라 필요 없는 컬럼 제거
			const cleanedData = jsonData.map(row => {
				if (type === "onch") {
					const {
						온채널주문코드, 상품명, 상품코드, 옵션, 수량, 가격
					} = row;
					return { 온채널주문코드, 온채널상품명: 상품명, 온채널상품코드: 상품코드, 옵션, 수량, 가격 };
				} else if (type === "coupang") {
					const {
						주문번호, 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비
					} = row;
					return { 쿠팡주문번호: 주문번호, 쿠팡상품번호: 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비 };
				}
				return row;
			});

			setData(cleanedData);
		};
		reader.readAsArrayBuffer(file);
	};

	// 데이터 병합 및 계산 로직
	const mergeAndDownload = () => {
		const mergedData = onchData.map((onchRow) => {
			const coupangRow = coupangData.find((c) => c.쿠팡주문번호 === onchRow.온채널주문코드);
			if (!coupangRow) return null;

			const 판매금액 = coupangRow.판매금액 || 0;
			const 판매배송비 = coupangRow.판매배송비 || 0;
			const 수수료 = 판매금액 * 0.108;
			const 순이익 = (onchRow.가격 - 판매금액 - 수수료 + 판매배송비) * onchRow.수량;

			return {
				...onchRow,
				쿠팡주문번호: coupangRow.쿠팡주문번호,
				쿠팡상품번호: coupangRow.쿠팡상품번호,
				판매금액,
				판매배송비,
				취소금액: coupangRow.취소금액,
				취소배송비: coupangRow.취소배송비,
				예상수수료: 수수료,
				순이익,
			};
		}).filter(row => row);

		// 전체 합계 행 추가
		const summaryRow = mergedData.reduce((acc, row) => {
			Object.keys(row).forEach((key) => {
				if (typeof row[key] === "number") {
					acc[key] = (acc[key] || 0) + row[key];
				}
			});
			return acc;
		}, { 온채널주문코드: "총합계" });
		mergedData.push(summaryRow);

		// 엑셀 파일로 변환 및 다운로드
		const worksheet = XLSX.utils.json_to_sheet(mergedData);
		const workbook = XLSX.utils.book_new();
		XLSX.utils.book_append_sheet(workbook, worksheet, "MergedData");
		XLSX.writeFile(workbook, "MergedData.xlsx");
	};

	return (
		<div>
			<h1>Excel Merge App</h1>
			<p>온채널 엑셀 파일 업로드:</p>
			<input
				type="file"
				accept=".xlsx, .xls"
				onChange={(e) => handleFileUpload(e, setOnchData, "onch")}
			/>
			<p>쿠팡 엑셀 파일 업로드:</p>
			<input
				type="file"
				accept=".xlsx, .xls"
				onChange={(e) => handleFileUpload(e, setCoupangData, "coupang")}
			/>
			<button onClick={mergeAndDownload}>Download Merged Excel</button>
		</div>
	);
}

export default ExcelMergeApp;


