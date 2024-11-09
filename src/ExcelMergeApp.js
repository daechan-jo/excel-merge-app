import React, { useState } from "react";
import * as XLSX from "xlsx";

function ExcelMergeApp() {
	const [onchData, setOnchData] = useState([]);
	const [coupangData, setCoupangData] = useState([]);

	const handleFileUpload = (e, setData, type) => {
		const file = e.target.files[0];
		const reader = new FileReader();
		reader.onload = (event) => {
			const data = new Uint8Array(event.target.result);
			const workbook = XLSX.read(data, { type: "array" });
			const sheetName = workbook.SheetNames[0];
			const sheet = workbook.Sheets[sheetName];
			const jsonData = XLSX.utils.sheet_to_json(sheet);

			const cleanedData = jsonData.map(row => {
				if (type === "onch") {
					const { 온채널주문코드, 상품명, 상품코드, 옵션, 수량, 가격 } = row;
					return { 온채널주문코드, 온채널상품명: 상품명, 온채널상품코드: 상품코드, 옵션, 수량, 가격 };
				} else if (type === "coupang") {
					const { 주문번호, 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비 } = row;
					return { 쿠팡주문번호: 주문번호, 쿠팡상품번호: 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비 };
				}
				return row;
			});

			setData(cleanedData);
		};
		reader.readAsArrayBuffer(file);
	};

	const mergeAndDownload = () => {
		const maxLength = Math.max(onchData.length, coupangData.length);
		const mergedData = [];

		for (let i = 0; i < maxLength; i++) {
			const { 온채널주문코드, 온채널상품명: 상품명, 온채널상품코드: 상품코드, 옵션, 수량, 가격 } = onchData[i] || {};
			const { 쿠팡주문번호: 주문번호, 쿠팡상품번호: 상품번호, 판매금액, 판매배송비, 취소금액, 취소배송비 } = coupangData[i] || {};

			const 수수료 = 판매금액 * 0.108;
			const 순이익 = (가격 - 판매금액 - 수수료 + 판매배송비) * (수량 || 1);
			const 판매 = 판매금액 + 판매배송비

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

		const summaryRow = mergedData.reduce((acc, row) => {
			Object.keys(row).forEach((key) => {
				if (typeof row[key] === "number" && key !== "쿠팡주문번호") {
					acc[key] = (acc[key] || 0) + row[key];
				}
			});
			return acc;
		}, { 온채널주문코드: "총합계" });
		mergedData.push(summaryRow);

		const worksheet = XLSX.utils.json_to_sheet(mergedData);

		// 스타일 적용
		const range = XLSX.utils.decode_range(worksheet["!ref"]);

		// 마지막 행(요약 행)에 스타일 추가
		for (let C = range.s.c; C <= range.e.c; ++C) {
			const cell = worksheet[XLSX.utils.encode_cell({ r: range.e.r, c: C })];
			if (cell) {
				cell.s = {
					fill: { fgColor: { rgb: "FFFF00" } },  // 노란색 배경
					font: { bold: true }                    // 굵은 글씨
				};
			}
		}

		// 마지막 컬럼에 스타일 적용
		for (let R = range.s.r; R <= range.e.r; ++R) {
			const cell = worksheet[XLSX.utils.encode_cell({ r: R, c: range.e.c })];
			if (cell) {
				cell.s = {
					fill: { fgColor: { rgb: "ADD8E6" } }    // 연한 파란색 배경
				};
			}
		}

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
