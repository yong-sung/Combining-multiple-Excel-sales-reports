import glob

판매보고들 = glob.glob(r"1.여러개의 엑셀 판매보고서 합치기\판매보고_*.xlsx") # 폴더 안에 있는 모든 판매보고_ 파일들을 리스트로 만들어줌
print(판매보고들)

for 판매보고 in 판매보고들:
    print(판매보고)