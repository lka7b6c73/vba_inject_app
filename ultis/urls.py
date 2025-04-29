def parse_form_data(form_data):
    """
    Xử lý dữ liệu từ form POST gửi lên
    """

    execute_url = form_data.get('execute_url')
    execute_path = form_data.get('execute_path')
    library_urls = form_data.getlist('library_urls[]')
    library_paths = form_data.getlist('library_paths[]')

    if not execute_path.endswith("\\") and not execute_path.endswith("/"):
        execute_path += "\\"

    # 1. url_list: file thực thi + các file thư viện
    url_list = [execute_url] + library_urls

    # 2. path_list: chỉ là thư mục chứa file thôi
    path_list = [execute_path] + library_paths

    # 3. main_file_path: thư mục + tên file từ URL chính
    filename = execute_url.split('/')[-1]
    main_file_path = execute_path + filename
    print(f"URL: {url_list}, PATH: {path_list}, MAIN_FILE_PATH: {main_file_path}")
    return url_list, path_list, main_file_path