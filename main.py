from PIL import Image, ImageDraw, ImageFont
from get_excel_details import get_photos , get_names_and_captions
import os 

count = 10
circle_size = (234,234)

frame_path = r'C:\Users\Risinu Wijesinghe\OneDrive\Desktop\Projects\Frame_Generator\template.jpg'
caption_font_path =r'C:\Users\Risinu Wijesinghe\OneDrive\Desktop\Projects\Frame_Generator\Oswald-ExtraLight.ttf'
name_font_path = r'C:\Users\Risinu Wijesinghe\OneDrive\Desktop\Projects\Frame_Generator\Oswald-ExtraLight.ttf'
excel_path =r"C:\Users\Risinu Wijesinghe\Favorites\Downloads\submissions.xlsx"
final_path =r'C:\Users\Risinu Wijesinghe\OneDrive\Desktop\Projects\Frame_Generator\edited_images'

def get_the_dimensions(img):
    cw,ch = img.size
    if (cw<ch):
        nh=397
        nw = int(nh/ch*cw)
        ph=119
        pw=80 + (397-nw)//2
    else:
        nw=397
        nh = int(nw/cw*ch) 
        pw =80
        ph =119 + (397-nh)//2

    return (nw,nh),(pw,ph)

def resize_image(image, target_size):
    return image.resize(target_size)

def create_circle_mask(size):
    mask = Image.new('L', size, 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0) + size, fill=255)
    return mask

def apply_circle_mask(image, mask):
    result = Image.new('RGBA', image.size, (244, 237, 231, 0))
    result.paste(image, mask=mask)
    return result

def add_name(image,name,path):
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(path,30)
    draw.text((580+(234-len(name)*10)//2,590),name,font=font,fill=(0,0,0))

def split_string(text, max_length):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        if len(current_line + word + ' ') <= max_length:
            current_line += word +' '
        else:
            lines.append(current_line)
            current_line = word + ' '

    return lines

def add_caption(image,caption,path):
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(path,20)
    if len(caption)>78:
        lineList=split_string(caption,60)
        for i in range(len(lineList)):
            draw.text((520,150+(30*i)),lineList[i],font=font,fill=(0,0,0))
    else:
        draw.text((500,150),caption,font=font,fill=(0,0,0))

contest_images,participant_images,err_nums=get_photos(excel_path,count)
names,captions = get_names_and_captions(excel_path,count)


for i in range(count + len(err_nums)):
    frame = Image.open(frame_path)
    curr_contest_img = contest_images[i]
    curr_part_img = participant_images[i]

    contest_dimension,paste_dimensions = get_the_dimensions(curr_contest_img)
    curr_contest_img = curr_contest_img.resize(contest_dimension)

    resized_image = resize_image(curr_part_img, circle_size)
    circle_mask = create_circle_mask(circle_size)
    circular_image = apply_circle_mask(resized_image, circle_mask)

    frame.paste(curr_contest_img,paste_dimensions)
    frame.paste(circular_image,(580,351))

    add_name(frame,names[i],name_font_path)
    add_caption(frame,captions[i],caption_font_path)

    frame.show()
    frame.save(os.path.join(final_path, f'{i+1} image.jpg'))
    





