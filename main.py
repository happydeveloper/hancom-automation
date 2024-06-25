import os
import win32com.client as win32

def main():
    try:
        # 파일 경로 설정
        file_root = os.path.abspath(os.path.dirname(__file__))
        file_path = os.path.join(file_root, 'untitledss.hwpx')

        # 한/글 객체 생성
        hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")
        hwp.RegisterModule('FilePathCheckerDLL', 'FilePathCheckerModule')

        # 한/글 프로그램 창 표시 설정
        hwp.XHwpWindows.Item(0).Visible = True

        # 파일 열기 시도
        #hwp.Open(None, "HWP", "forceopen:true")

        # 삽입할 텍스트 생성 및 실행
        for i in range(10000):
            gun = hwp.CreateAction("InsertText")
            bullet = gun.CreateSet()
            gun.GetDefault(bullet)
            bullet.SetItem("Text", "한컴 자동화")
            gun.Execute(bullet)
            print(i)

        print(f"{file_path} 파일이 성공적으로 열렸습니다.")
        #hwp.Quit()
        
    except Exception as e:
        print(f"오류가 발생하여 파일을 열 수 없습니다: {e}")

if __name__ == "__main__":
    main()
