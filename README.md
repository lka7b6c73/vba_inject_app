Khánh ngày 29/4/2025 gửi Khánh của tương lai
Nếu mà bạn đã đọc được dòng này có nghĩa là tôi đã không con thông minh nữa mà bị ngu đi rồi nên mới vào đây để xem lại cách nó hoạt động.
Đây nha, code này có folder ultis chứa các file xử lý backend
1. file urls.py chỉ là lấy dữ liệu link url và path lưu lại xong sửa lại cho tương thích thôi
Dữ liệu sẽ có 2 phần 
    1 là excute file, đây là file thực thi 
    2 là library file, đây là file thư viện kiểu cần để cho file excute file chạy được
2. original.py là file sẽ lấy dữ liệu từ urls.py sau đó tạo ra vba subA và subB
    subA là cái download tất cả các link còn 
    subB là cái thực thi file thực thi á.
    sub A sẽ hoạt động theo kiểu lấy dữ liệu link là 1 hàm array và 1 array path lưu
    sau đó tải lần lượt về. mỗi vòng for sẽ có tải và kiểm tra xem tải thành công chưa nếu thành công thì mới tải file tiếp theo
    sub B thì chỉ đơn giản là gọi shell ra đê thực thi cho dễ thôi
3. base.py là file sẽ lấy dữ liệu vba subA Và B để từ đó tạo ra vba mà tôi sẽ inject
    vba này sẽ có subA end và subB end mục đích là tí tôi sẽ mã hóa và thay thế subA và B đó. Làm thế để không bị lỗi
    vba còn có hàm mã hóa này . có nghĩa sẽ láy vba sub A và sub B mã hóa nó đi và chèn vào vba. tiếp đó thì khi mà hàm chạy thì sẽ mã hóa code trước đó và xóa subA và subB cũ xong thay thế vào đó, sau 0,002s thì sẽ thực thi hàm A và B.
    sau khi tạo xong thì sẽ gủi đi
4. injector.py này là file sẽ chèn vba vào word vầ excel
5. uploads folder là folder chứa cái mà mình upload lên
6. folder generated là chứa các file đã chèn vba vào rồi

có thể tham khảo ở original.txt là code sub A và sub B
base.txt là code full sau mã hóa.

