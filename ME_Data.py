class ME:
    def __init__(self, me_id):
        self.me_id = me_id
        self.initialized = False


class Cell:
    def __init__(self, cell_id, sector_id, tac, pci, band_indicator, dl_sys_bw,
                 dl_center_freq, ul_sys_bw, ul_center_freq, root_seq_st_num):
        self.cell_id = cell_id
        # self.frequency = frequency
        self.sector_id = sector_id
        self.tac = tac
        self.pci = pci
        self.band_indicator = band_indicator
        self.dl_sys_bw = dl_sys_bw
        self.dl_center_freq = dl_center_freq
        self.ul_sys_bw = ul_sys_bw
        self.ul_center_freq = ul_center_freq
        self.root_seq_st_num = root_seq_st_num

    def __init__(self, cell_id, sector_id, tac, pci):
        self.cell_id = cell_id
        # self.frequency = frequency
        self.sector_id = sector_id
        self.tac = tac
        self.pci = pci
        # self.band_indicator = band_indicator
        # self.dl_sys_bw = dl_sys_bw
        # self.dl_center_freq = dl_center_freq
        # self.ul_sys_bw = ul_sys_bw
        # self.ul_center_freq = ul_center_freq
        # self.root_seq_st_num = root_seq_st_num
