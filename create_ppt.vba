Sub CreateSemiconductorPackagingPresentation()
    Dim pptApp As Object
    Dim pptPres As Object
    Dim slideIndex As Integer
    Dim slide As Object
    Dim titleLayout As Object
    Dim contentLayout As Object
    
    ' PowerPoint 애플리케이션 생성
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set pptApp = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add
    
    ' 레이아웃 설정 (1: Title, 2: Text)
    
    ' --- Slide 1: Title ---
    Set slide = pptPres.Slides.Add(1, 1) ' ppLayoutTitle
    slide.Shapes(1).TextFrame.TextRange.Text = "반도체 패키징의 대전환"
    slide.Shapes(2).TextFrame.TextRange.Text = "소재가 주도하는 More than Moore 시대" & vbCrLf & "2026 세미나 강의 교안"
    
    ' --- Slide 2: 강의 개요 ---
    Set slide = pptPres.Slides.Add(2, 2) ' ppLayoutText
    slide.Shapes(1).TextFrame.TextRange.Text = "강의 개요"
    slide.Shapes(2).TextFrame.TextRange.Text = "주제: 반도체 패키징의 본질적 정의, 세대적 진화, AI 시대의 소재 혁신" & vbCrLf & _
                                               "목표: 첨단 패키징 기술(2.5D/3D, HBM) 이해 및 핵심 소재 트렌드 파악"
                                               
    ' --- Slide 3: 패키징의 4대 핵심 기능 ---
    Set slide = pptPres.Slides.Add(3, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "패키징의 4대 핵심 기능"
    slide.Shapes(2).TextFrame.TextRange.Text = "1. 보호(Protection): 습기, 먼지, 충격으로부터 칩 보호" & vbCrLf & _
                                               "2. 연결(Electrical Connection): 칩(나노) ↔ PCB(밀리) 간 배선 연결 (RDL, Bump)" & vbCrLf & _
                                               "3. 열 관리(Thermal Management): 1000W+ 발열 해소, 수명/속도 보장" & vbCrLf & _
                                               "4. 무결성(Signal & Power Integrity): 신호 간섭 최소화, 안정적 전력 공급"

    ' --- Slide 4: 패키징 기술의 역사 ---
    Set slide = pptPres.Slides.Add(4, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "패키징 기술의 역사 (1세대 ~ 4세대)"
    slide.Shapes(2).TextFrame.TextRange.Text = "1세대 (Leadframe): DIP, QFP (와이어 본딩) - 단순, 저렴" & vbCrLf & _
                                               "2세대 (Substrate): BGA (솔더 볼) - I/O 증가" & vbCrLf & _
                                               "3세대 (Miniaturization): CSP, WLP - 칩 크기 초소형화" & vbCrLf & _
                                               "4세대 (Advanced): TSV, Fan-out, 2.5D/3D - 이종 집적, AI 시대 핵심"

    ' --- Slide 5: AI 시대와 '메모리 벽' ---
    Set slide = pptPres.Slides.Add(5, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "AI 시대와 '메모리 벽'"
    slide.Shapes(2).TextFrame.TextRange.Text = "문제: GPU 연산 속도 >> 메모리 전송 속도 (병목 현상)" & vbCrLf & _
                                               "해결: 칩 간 거리 최소화 (Micro-meter 단위) → Advanced Packaging" & vbCrLf & _
                                               "도전: 전력 밀도 급증 (700W → 1500W) → 방열 설계 필수"

    ' --- Slide 6: 소재 혁신 1 - HBM & Hybrid Bonding ---
    Set slide = pptPres.Slides.Add(6, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "소재 혁신 1: HBM & Hybrid Bonding"
    slide.Shapes(2).TextFrame.TextRange.Text = "Trend: 범프(Bump)가 사라진다 (Bump-less)" & vbCrLf & _
                                               "소재: Cu-Cu 직접 접합 (Hybrid Bonding)" & vbCrLf & _
                                               "효과: 인터커넥트 밀도 100배↑, 대역폭 극대화, Latency 단축"

    ' --- Slide 7: 소재 혁신 2 - RDL & 저유전 소재 ---
    Set slide = pptPres.Slides.Add(7, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "소재 혁신 2: RDL & 저유전 소재"
    slide.Shapes(2).TextFrame.TextRange.Text = "Trend: 고속 신호 손실 최소화" & vbCrLf & _
                                               "소재: 에폭시 → PID (감광성 절연 소재, PPE/PI)" & vbCrLf & _
                                               "특성: Low-Dk (<2.5), Low-Df (<0.001)" & vbCrLf & _
                                               "효과: 신호 손실 40% 절감, 초미세 회로 구현"

    ' --- Slide 8: 소재 혁신 3 - 유리 기판 ---
    Set slide = pptPres.Slides.Add(8, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "소재 혁신 3: 유리 기판 (Glass Substrate)"
    slide.Shapes(2).TextFrame.TextRange.Text = "Trend: 대면적 패키징의 '휨(Warpage)' 해결" & vbCrLf & _
                                               "소재: 유기 기판 → 유리 기판 (Si와 유사한 CTE ~3ppm/K)" & vbCrLf & _
                                               "효과: 휨 50%↓, 대형화 가능, TGV 이용한 직접 전송"

    ' --- Slide 9: 소재 혁신 4 - TIM (열 관리) ---
    Set slide = pptPres.Slides.Add(9, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "소재 혁신 4: TIM (열 관리)"
    slide.Shapes(2).TextFrame.TextRange.Text = "Trend: 극한의 발열 제어" & vbCrLf & _
                                               "소재: 그리스 → 인듐, 액체금속, PCM (상변화 소재)" & vbCrLf & _
                                               "효과: 정션 온도 80°C 이하 유지, 스로틀링 방지"

    ' --- Slide 10: 미래 과제 및 결론 ---
    Set slide = pptPres.Slides.Add(10, 2)
    slide.Shapes(1).TextFrame.TextRange.Text = "미래 과제 및 결론"
    slide.Shapes(2).TextFrame.TextRange.Text = "전략 1: Purity Control (10ppb 이하, Cu 산화 방지)" & vbCrLf & _
                                               "전략 2: PFAS-Free (환경 규제 대응)" & vbCrLf & _
                                               "전략 3: Digital Twin (시뮬레이션 기반 소재 개발)" & vbCrLf & vbCrLf & _
                                               "결론: 원료 업체는 단순 공급자가 아닌 '성능 설계자(Performance Architect)'가 되어야 합니다."

    MsgBox "프레젠테이션 생성이 완료되었습니다!", vbInformation, "완료"
End Sub
