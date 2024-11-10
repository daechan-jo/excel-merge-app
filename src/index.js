import React from "react";
import ReactDOM from "react-dom";
import { HashRouter } from "react-router-dom";
import ExcelMergeApp from "./ExcelMergeApp";

ReactDOM.render(
	<HashRouter>
		<ExcelMergeApp />
	</HashRouter>,
	document.getElementById("root")
);
