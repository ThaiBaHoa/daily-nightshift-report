// Field instructions for different STT values
// These provide detailed guidance for each inspection section

export const FIELD_INSTRUCTIONS: { [key: string]: string } = {
  "1": `Chi tiết kiểm tra: 
    a. Quan sát việc tuân thủ kiểm tra FOD bãi đậu trước khi tàu đáp và sau khi đưa tàu đi của NVKT.
    b. Croscheck tại các bãi có bảo dưỡng lớn xem team bải dưỡng có tuân thủ quy trình ngăn ngừa FOD hay không?`,

  "2": `Chi tiết kiểm tra:
    Random check các tàu thực hiện WKLY có được drain fuel theo WKLY checklist hay không?`,

  "3": `Có thể kiểm tra OWP feedback của ngày hôm trước, với các tiêu chí:
    a. Các WO từ chối phải có lí do rõ ràng và được sự xác nhận từ MOC.
    b. Thông tin daily check phải được điền đầy đủ.
    c. Các WO được thực hiện phải ghi nhận số chứng chỉ của NVKT rõ ràng.`,

  "6": `Chi tiết kiểm tra:
    a. Kiểm tra các dấu hiệu mục của buồng hàng - đặc biệt là khu vực quanh mép buồng hàng và cửa buồng hàng.
    b. Kiểm tra tình trạng TDP.
    c. Kiểm tra các tầm linning buồng hàng.`,

  "7": `Chi tiết kiểm tra:
    a. Số lượng hóa chất.
    b. Số lượng lần rửa.`,

  "8": `Chi tiết kiểm tra:Tăng cường random check áo phao`,

  "11": `Chi tiết kiểm tra: Tăng cường kiểm tra vị trí Anti-ice access panel`,

  "19": `Chi tiết kiểm tra: PTS là chương trình yêu cầu các đơn vị phải hoàn thành phần việc của mình trong khoảng thời gian CỐ ĐỊNH và TỐI ƯU NHẤT đã được thống nhất bằng văn bản. Nhằm duy trì tổng thời gian dừng/nghỉ giữa các chuyến bay TỐI ƯU.`
};

// Helper function to get instruction for a specific STT
export const getFieldInstruction = (stt: string | number): string => {
  const sttKey = String(stt);
  return FIELD_INSTRUCTIONS[sttKey] || "Không có hướng dẫn chi tiết cho mục kiểm tra này.";
};
