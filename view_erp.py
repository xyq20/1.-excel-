from chrome_erp_session import ChromeErpSession


def main() -> None:
    with ChromeErpSession() as session:
        session.keep_open()
    print("ERP专用Chrome已打开；此操作不会抓取数据或修改Excel。")


if __name__ == "__main__":
    main()
