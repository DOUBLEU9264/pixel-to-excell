package main

import (
	"bytes"
	"errors"
	"fmt"
	"image"
	"image/jpeg"
	"image/png"
	"log"
	"os"
	"path"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	filepath := args()
	img, err := readimg(filepath)
	if err != nil {
		log.Println(err)
		os.Exit(1)
	}

	f := excelize.NewFile()
	defer f.Close()

	sheetName := filenameWithoutExt(filepath)
	err = f.SetSheetName("Sheet1", sheetName)
	if err != nil {
		log.Println(err)
		os.Exit(1)
	}

	err = proc(img, f, sheetName)
	if err != nil {
		log.Println(err)
		os.Exit(1)
	}

	f.SaveAs("a.xlsx")
	fmt.Println("Done.")
}

func readimg(filepath string) (image.Image, error) {
	ext := path.Ext(filepath)
	if ext == "" {
		return nil, errors.New("unknown file type")
	}

	var buf []byte
	var err error
	switch ext {
	case ".jpg", ".JPG", ".jpeg", ".JPEG":
		buf, err = os.ReadFile(filepath)
		if err != nil {
			return nil, errors.Join(fmt.Errorf("open file: %s error", filepath), err)
		}
		img, err := jpeg.Decode(bytes.NewReader(buf))
		if err != nil {
			return img, errors.Join(fmt.Errorf("decode pic %s error", filepath), err)
		}
		return img, nil
	case ".png", ".PNG":
		buf, err = os.ReadFile(filepath)
		if err != nil {
			return nil, errors.Join(fmt.Errorf("open file: %s error", filepath), err)
		}
		img, err := png.Decode(bytes.NewReader(buf))
		if err != nil {
			return img, errors.Join(fmt.Errorf("decode pic %s error", filepath), err)
		}
		return img, nil
	default:
		return nil, errors.New("unsupported file type")
	}
}

func filenameWithoutExt(filepath string) string {
	filename := path.Base(filepath)
	return strings.TrimSuffix(filename, path.Ext(filename))
}

func args() string {
	args := os.Args
	if len(args) < 2 {
		fmt.Println("Usage: ./pixel-to-excell $path-to-image")
		os.Exit(1)
	}

	return args[1]
}

func proc(img image.Image, f *excelize.File, sheetName string) error {
	imgWidth, imgHeight := picSize(img)
	if imgWidth == 0 || imgHeight == 0 {
		return errors.New("blank file")
	}

	fmt.Printf("pic width: %dpx, height: %dpx\n", imgWidth, imgHeight)

	widthCellName, _ := excelize.CoordinatesToCellName(imgWidth, 1)
	f.SetColWidth(sheetName, "A", strings.TrimSuffix(widthCellName, "1"), 2.85)
	style := excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Color: []string{""}, Pattern: 1},
	}

	for row := 0; row < imgHeight; row++ {
		fmt.Printf("\rProcessing row no.%d ...", row+1)
		for column := 0; column < imgWidth; column++ {
			c := img.At(column, row)
			hexColor := rgbaToHexRGB(c.RGBA())
			style.Fill.Color[0] = hexColor
			styleID, _ := f.NewStyle(&style)
			cellName, _ := excelize.CoordinatesToCellName(column+1, row+1)
			f.SetCellStyle(sheetName, cellName, cellName, styleID)
		}
	}

	fmt.Println()

	return nil
}

func picSize(img image.Image) (width, height int) {
	size := img.Bounds().Size()
	width = size.X
	height = size.Y
	return
}

func rgbaToHexRGB(cr, cg, cb, caRAW uint32) string {
	ca := float64(caRAW / 0xffff)
	r := uint32(float64(cr) / 0xffff * 0xff * ca)
	g := uint32(float64(cg) / 0xffff * 0xff * ca)
	b := uint32(float64(cb) / 0xffff * 0xff * ca)
	return fmt.Sprintf("%02X%02X%02X", r, g, b)
}
