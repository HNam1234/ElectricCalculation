# ElectricCalculation – Hướng dẫn sử dụng

## Mục tiêu

ElectricCalculation là ứng dụng Windows hỗ trợ quản lý chỉ số điện theo kỳ (tháng/năm): nhập danh sách khách hàng, nhập chỉ số công tơ, tự tính kWh/tiền điện, lưu bộ dữ liệu để mở lại và xuất Excel/PDF khi cần.

---

## Quy trình sử dụng (khuyến nghị)

1. Mở ứng dụng → vào màn hình **Startup**
2. Chọn **Create dataset** → **Import from Excel**
3. Hoàn tất **Import wizard** → Import dữ liệu
4. Ở màn hình chính, nhập/sửa `CurrentIndex` (và các trường liên quan)
5. Chọn **Save Snapshot** để lưu lại → **Export** khi cần in/xuất

---

## Startup (quản lý bộ dữ liệu)

Tại Startup, bạn có thể:

- **Open dataset**: mở snapshot đã lưu để chỉnh sửa tiếp
- **Pin / Unpin**: ghim/bỏ ghim dataset quan trọng
- **Delete**: xoá snapshot
- **Open snapshot folder**: mở thư mục snapshot trong Documents
- **Create dataset**:
  - **Manual entry**: tạo danh sách trống và nhập tay
  - **Import from Excel**: nhập từ Excel `.xlsx` (khuyến nghị)
- **Create new period from dataset**: tạo kỳ mới dựa trên dataset cũ (carry forward chỉ số)

---

## Nhập từ Excel (.xlsx)

Import wizard gồm 3 bước:

1. **Chọn file + sheet**: chọn file `.xlsx` và sheet cần nhập (hỗ trợ kéo-thả)
2. **Chọn dòng header**: chọn dòng chứa tên cột (wizard có preview để kiểm tra)
3. **Map cột Excel ↔ field**: kiểm tra mapping gợi ý và chỉnh lại nếu cần

Sau khi import, danh sách khách hàng sẽ hiển thị trong màn hình chính.

---

## Màn hình chính (nhập chỉ số / tính tiền)

### Trường dữ liệu thường dùng

- `PreviousIndex`: chỉ số cũ
- `CurrentIndex`: chỉ số mới
- `Multiplier`: hệ số
- `SubsidizedKwh`: bao cấp (kWh miễn/giảm)
- `UnitPrice`: đơn giá

Ứng dụng tự tính kWh và thành tiền, đồng thời hiển thị trạng thái (OK/thiếu/lỗi/cảnh báo) để bạn xử lý nhanh.

### Chế độ nhập liệu

- **Fast Entry**: tối ưu nhập nhanh trong DataGrid (phù hợp nhập nhiều `CurrentIndex`)
- **Detail**: hiển thị thêm thông tin chi tiết

---

## Lưu dữ liệu

- **Save Data File (.json)**: lưu ra file `.json` ở vị trí bạn chọn (phù hợp sao lưu theo thư mục riêng).
- **Save Snapshot**: lưu nhanh vào thư mục snapshot để lần sau mở lại từ Startup.

---

## Xuất dữ liệu

- **Xuất Excel tổng hợp**: xuất theo template tổng hợp.
- **Xuất hoá đơn Excel**: xuất theo template hoá đơn (1 khách hoặc nhiều khách).
- **Xuất PDF**: yêu cầu máy có Microsoft Excel (ứng dụng dùng Excel để export PDF).

---

## Phím tắt nhập liệu (DataGrid)

- `Enter`: commit ô hiện tại và chuyển sang dòng kế tiếp (cùng cột)
- `Ctrl+V`: dán dữ liệu dạng bảng từ clipboard (tương tự dán từ Excel)
- `Ctrl+D`: fill-down (copy giá trị xuống các dòng đã chọn)
- `Ctrl+Shift+D`: nhân bản dòng đang chọn
- `Delete`: xoá các dòng đang chọn
- `Ctrl+F`: đưa con trỏ vào ô tìm kiếm

---

## Settings (giá trị mặc định)

Settings cho phép đặt mặc định khi thêm dòng mới hoặc khi import:

- `DefaultUnitPrice` (đơn giá)
- `DefaultMultiplier` (hệ số)
- `DefaultSubsidizedKwh` (bao cấp)
- `DefaultPerformedBy` (người ghi)

---

## File nằm ở đâu?

Mặc định ứng dụng lưu trong Documents (My Documents):

- Settings: `Documents/ElectricCalculation/settings.json`
- Import mapping profiles: `Documents/ElectricCalculation/import_mapping_profiles.json`
- Pins: `Documents/ElectricCalculation/pinned_datasets.json`
- Snapshots: `Documents/ElectricCalculation/Saves/`

Template trong repo (source):

- Hoá đơn: `DefaultTemplate.xlsx`
- Tổng hợp/sample: file `.xlsx` nằm cạnh `ElectricCalculation/ElectricCalculation.sln`

---

## Thiết kế tổng quan (MVVM) – High level design

### Nguyên tắc

- **View (XAML)**: chỉ hiển thị và binding.
- **ViewModel**: state + command (`RelayCommand`) để View gọi hành động.
- **Model/Service**: dữ liệu và xử lý import/export/lưu file.
- Code-behind của View chỉ nên là “view glue” (không chứa business logic).

### Cấu trúc thư mục

- View: `ElectricCalculation/ElectricCalculation/Views/`
- ViewModel: `ElectricCalculation/ElectricCalculation/ViewModels/`
- View glue: `ElectricCalculation/ElectricCalculation/Behaviors/`, `Converters/`, `Helpers/`
- Model: `ElectricCalculation/ElectricCalculation/Models/`
- Service: `ElectricCalculation/ElectricCalculation/Services/`

```mermaid
flowchart LR
  V[Views (XAML)] --> VM[ViewModels (Commands/State)]
  V --> Glue[Behaviors/Converters]
  VM --> M[Models]
  VM --> S[Services]
  S --> Files[(.xlsx / .json / .pdf)]
```
