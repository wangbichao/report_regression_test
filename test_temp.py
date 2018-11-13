import io
import sys


h_total = 4399
v_total = 2249
refresh_rate = 120
bytes_per_pixel = 8
v_scale_ratio = 1.5
vp_width = 3840
vp_height = 2160
pixel_clock = 1188000000

refresh_rate = pixel_clock / h_total / v_total  
line_period = h_total / pixel_clock
frame_period = h_total * v_total / pixel_clock

bandwidth_w = vp_width / (h_total / pixel_clock) * bytes_per_pixel * v_scale_ratio

print(bandwidth_w/(1E9))

bandwidth_w_h = vp_width * vp_height / frame_period * bytes_per_pixel * v_scale_ratio

print(bandwidth_w_h/(1E9))

print(3748-318-5)




