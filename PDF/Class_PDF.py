 # Add the logo at the top-left corner
        self.image('C:\\Users\\phath\\OneDrive\\Hình ảnh\\download.jpeg', 10, 10, 20, 20)  # 2cm x 2cm (20mm x 20mm)



pdf.add_font('Times', '', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\Times\\TIMES.TTF', uni=True)  # Add normal Times New Roman
pdf.add_font('Times', 'B', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\Times\\TIMESBD.TTF', uni=True)  # Add bold version
pdf.add_font('Times', 'I', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\Times\\TIMESI.TTF', uni=True)  # Add italic version
pdf.add_font('Times', 'BI', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\Times\\TIMESBI.TTF', uni=True)  # Add bold-italic version


# ** Add the custom Unicode font (DejaVu) **:
pdf.add_font('DejaVu', '', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\dejavu-sans\\ttf\\DejaVuSans.ttf', uni=True)  # Add normal DejaVu
pdf.add_font('DejaVu', 'B', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\dejavu-sans\\ttf\\DejaVuSans-Bold.ttf', uni=True)  # Add bold version
pdf.add_font('DejaVu', 'I', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\dejavu-sans\\ttf\\DejaVuSans-Oblique.ttf', uni=True)  # Add italic version
pdf.add_font('DejaVu', 'BI', 'D:\\DuAnPython\\SmartBees_PupilsManagement_ver2.0\\PDF\\dejavu-sans\\ttf\\DejaVuSans-BoldOblique.ttf', uni=True)  # Add bold-italic version
