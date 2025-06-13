import { ExcelManager } from "./excel/excel-manager";

const button = document.getElementById("button-handler") as HTMLButtonElement;

const column = [
    {
        header: "Thông tin khách hàng",
        children: [
            { key: "name", header: "Tên" },
            { key: "age", header: "Tuổi" }
        ]
    },
    {
        header: "Địa chỉ",
        key: "address"
    }
];

const data = [
    { name: "Nguyễn Văn A", age: 25, address: "Hà Nội" },
    { name: "Trần Thị B", age: 30, address: "Đà Nẵng" },
    { name: "Lê Văn C", age: 22, address: "Hồ Chí Minh" }
];

const titleList = [
    {
        content: "BÁO CÁO DANH SÁCH KHÁCH HÀNG",
        bold: true,
        fontSize: 16,
        width: 7
    }
];


button.addEventListener("click", async () => {
    const buffer = await ExcelManager.load(column, data, titleList)
    ExcelManager.downloadBuffer(buffer, "demo.xlsx")
})