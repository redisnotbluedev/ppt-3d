from PIL import Image, ImageDraw

vertices = [(-1,-1,9),(1,-1,9),(1,1,9),(-1,1,9),(-1,-1,11),(1,-1,11),(1,1,11),(-1,1,11)]
edges = [(0,1),(1,2),(2,3),(3,0),(4,5),(5,6),(6,7),(7,4),(0,4),(1,5),(2,6),(3,7)]
camera = (0, 0, 0)
focal_distance = 2000
width = 600
height = 600
points = []

radius = 3
image = Image.new("L", (width, height), color="white")
draw = ImageDraw.Draw(image)

for vertex in vertices:
	if vertex[2] <= camera[2]:
		continue

	x = focal_distance * vertex[0] / vertex[2]
	y = focal_distance * vertex[1] / vertex[2]
	screen_x = x + width / 2
	screen_y = -y + height / 2

	#draw.ellipse((screen_x - radius, screen_y - radius, screen_x + radius, screen_y + radius), outline="black", fill="black")
	points.append((screen_x, screen_y))

for edge in edges:
	start = points[edge[0]]
	end = points[edge[1]]
	draw.line((start, end), width=2)

image.show()