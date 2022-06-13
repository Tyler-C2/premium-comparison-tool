import cv2
import pytesseract
from pytesseract import Output

class MainImage():
	def __init__(self):
		self.path = None
		self.file_name = None
		self.image = None
		self.height = None
		self.width = None

	def add_path(self, file_path):
		supported = ["png","jpg","jpeg"]

		if file_path[-3:] in supported or file_path[-4:] in supported:
			self.path = file_path
			self.load_img()
		else:
			self.path == self.path

	def load_img(self):
		self.image = cv2.imread(self.path)
		h_and_w = self.get_height_and_width(self.image)
		self.height, self.width = h_and_w[0], h_and_w[1]

	def get_height_and_width(self, img):
		height, width = img.shape[:2]
		return (height, width)

	def get_image_data(self, img):
			return pytesseract.image_to_data(img, output_type=Output.DICT, config=r'--oem 3 --psm 4', lang='eng')

	def resize_img(self, origin_image, new_height):
		h_and_w = self.get_height_and_width(origin_image)
		ratio = new_height/float(h_and_w[0])
		new_dimension = (int(ratio*h_and_w[1]), new_height)

		resizedImage = cv2.resize(origin_image, new_dimension, interpolation = cv2.INTER_LANCZOS4)
		
		return resizedImage

class Top_ROI(MainImage):
	def __init__(self):
		super().__init__()
		self.top_roi_img = None
		self.data = None
		self.vehicle_text = ""

	def creator(self):
		self.get_start_region()

		self.data = self.get_image_data(self.top_roi_img)

		self.get_vehicle()	

	def clear(self):
		self.vehicle_text = ""

	def get_start_region(self):
		self.top_roi_img =  self.resize_img(self.image[0:int(self.height//5),0:int(self.width//1.5)], 900)

	def get_vehicle(self):
		for i in range(len(self.data['text'])):
			if len(self.data['text'][i]) == 4 and self.data['text'][i].isdigit():
				self.vehicle_text = f"{self.data['text'][i]} {self.data['text'][i+1]} {self.data['text'][i+2]}"
				break

		self.vehicle_text = self.vehicle_text.strip()

class Right_ROI(MainImage):
	def __init__(self):
		super().__init__()
		self.right_roi_img = None
		self.data = None
		self.size_of_col=None
		self.premium_col=None
		self.premium_values = []

	def creator(self):
		self.get_start_region()

		self.data = self.get_image_data(self.right_roi_img)

		self.get_premium_col()
		self.get_premium_col_data()

	def clear(self):
		self.path = None
		self.premium_values = []

	def get_start_region(self):
		self.right_roi_img = self.resize_img(self.image[int(self.height//5):self.height,int(self.width//2):self.width], 900)

	def get_premium_col(self):

		self.find_premium_bounds()

		temp = self.right_roi_img[
			self.size_of_col[1] : (self.size_of_col[1] + self.size_of_col[2]),
			self.size_of_col[0] : (self.size_of_col[0] + self.size_of_col[3])
		]

		self.premium_col = self.resize_img(temp,900)

	def find_premium_bounds(self):

		size_of_data = len(self.data['level'])
		idx_of_last_occurance = 0

		for i in range(size_of_data):
			if self.data['text'][i] == 'Premium':
				start_idx = i

		end_of_col = int(self.right_roi_img.shape[1]) 

		self.get_premium_col_size(start_idx, end_of_col)

	def get_premium_col_size(self, start_of_col, end_point):
	    x = self.data['left'][start_of_col] - 10
	    y = self.data['top'][start_of_col] - 10
	    h = end_point
	    w = (self.data['width'][start_of_col]*2)+10

	    self.size_of_col = (x, y, h, w)

	def get_premium_col_data(self):
		img_blur = cv2.GaussianBlur(self.premium_col, (3,3),0)

		data = self.get_image_data(img_blur)

		for i in range(len(data['level'])):
			if data['text'][i] != 'Premium' and data['text'][i] != '$' and data['text'][i] != '':
				self.premium_values.append(data['text'][i])
		
		self.premium_values.insert(3,'N/A')
	
