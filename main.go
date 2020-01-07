package main

import (
	"encoding/json"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"net/http"
	"strings"
	"sync"
)

func main() {
	fmt.Println("编辑 vedios.xlsx 表格，并关闭表格，按 回车 键开始读取...")
	
	var in string
	fmt.Scanln(&in)
	
	fmt.Println("视频号\t点赞数\t投币数\t收藏数")
	
	// 读取文件，判断文件是否存在
	f, err := excelize.OpenFile("vedios.xlsx")
	if err != nil {
		fmt.Println(err.Error())
		return
	}

	var wg sync.WaitGroup
	readQueue := make(chan *model, 100)
	writeQueue := make(chan *model, 100)

	for i := 0; i < 5; i++ {
		wg.Add(1)
		go work(readQueue, writeQueue, &wg)
	}

	// 读取视频ID
	rows := f.GetRows("Sheet1")
	if len(rows) == 0 {
		fmt.Println("Excel文件错误，检查文件！")
		return
	}
	rows = rows[1:]
	for i, row := range rows {
		readQueue <- &model{
			Line: i + 2,
			ID:   strings.ReplaceAll(row[0], "https://www.bilibili.com/video/av", ""),
		}
	}
	// 表格读取完毕，关闭读取队列
	close(readQueue)

	go func() {
		for {
			row, ok := <-writeQueue
			if !ok {
				break
			}
			f.SetCellValue("Sheet1", fmt.Sprintf("B%v", row.Line), row.Info.Data.AID)
			f.SetCellValue("Sheet1", fmt.Sprintf("C%v", row.Line), row.Info.Data.Like)
			f.SetCellValue("Sheet1", fmt.Sprintf("D%v", row.Line), row.Info.Data.Coin)
			f.SetCellValue("Sheet1", fmt.Sprintf("E%v", row.Line), row.Info.Data.Favorite)
			fmt.Println(fmt.Sprintf("av%v", row.Info.Data.AID), row.Info.Data.Like, row.Info.Data.Coin, row.Info.Data.Favorite)
		}
	}()

	// 等待统计完成，关闭写通道
	wg.Wait()
	close(writeQueue)
	// 保存文件
	if err = f.Save(); err != nil {
		fmt.Println("表格保存失败，请确认表格未打开，",err.Error())
	} else {
		fmt.Println("读取完成，数据已保存！")
	}
	fmt.Println()
	fmt.Println("按 回车 键退出...")
	fmt.Scanln(&in)
}

func work(readQueue <-chan *model, writeQueue chan<- *model, wg *sync.WaitGroup) {
	defer wg.Done()
	for {
		row, ok := <-readQueue
		if !ok {
			return
		}
		info, _ := getInfo(row.ID)
		row.Info = info
		writeQueue <- row
	}
}

func getInfo(id string) (*avInfo, error) {
	url := fmt.Sprintf("https://api.bilibili.com/x/web-interface/archive/stat?aid=%v", id)
	resp, err := http.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()
	info := &avInfo{}
	err = json.NewDecoder(resp.Body).Decode(info)
	if err != nil {
		return nil, err
	}
	return info, nil
}

type model struct {
	Line int
	ID   string
	Info *avInfo
}

type avInfo struct {
	Code    int    `json:"code"`
	Message string `json:"message"`
	Data    struct {
		AID      int `json:"aid"`
		Like     int `json:"like"`
		Coin     int `json:"coin"`
		Favorite int `json:"favorite"`
	} `json:"data"`
}
