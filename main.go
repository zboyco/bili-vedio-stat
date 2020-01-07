package main

import (
	"encoding/json"
	"errors"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"net/http"
	"strings"
	"sync"
	"time"
)

func main() {
	fmt.Println("编辑 vedios.xlsx 表格，并关闭表格，按 回车 键开始读取...")

	var in string
	fmt.Scanln(&in)

	defer func() {
		fmt.Println()
		fmt.Println("按 回车 键退出...")
		fmt.Scanln(&in)
	}()

	// 读取文件，判断文件是否存在
	f, err := excelize.OpenFile("vedios.xlsx")
	if err != nil {
		fmt.Println(err.Error())
		return
	}

	// 读取视频地址
	rows := f.GetRows("Sheet1")
	if len(rows) == 0 {
		fmt.Println("Excel文件错误，检查表格！")
		return
	}

	var wg sync.WaitGroup
	readQueue := make(chan *model, 100)
	writeQueue := make(chan *model, 100)

	for i := 0; i < 3; i++ {
		wg.Add(1)
		go work(readQueue, writeQueue, &wg)
	}

	// 读取统计数据
	rows = rows[1:]
	for i, row := range rows {
		if strings.Contains(row[0], "https://www.bilibili.com/video/av") {
			readQueue <- &model{
				Line: i + 2,
				ID:   strings.ReplaceAll(row[0], "https://www.bilibili.com/video/av", ""),
			}
		}
	}
	// 读取完毕，关闭读取队列
	close(readQueue)

	fmt.Println("视频号\t点赞数\t投币数\t收藏数")
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

	// 等待统计写入完成
	wg.Wait()
	// 关闭写通道
	close(writeQueue)
	// 保存文件
	if err = f.Save(); err != nil {
		fmt.Println("表格保存失败，请确认表格处于未打开，", err.Error())
	} else {
		fmt.Println("读取完成，数据已保存！")
	}
}

func work(readQueue <-chan *model, writeQueue chan<- *model, wg *sync.WaitGroup) {
	defer wg.Done()
	for {
		row, ok := <-readQueue
		if !ok {
			return
		}
		info, err := getInfo(row.ID)
		if err != nil {
			fmt.Printf("av%v 错误：%v", row.ID, err)
			continue
		}
		row.Info = info
		writeQueue <- row
		time.Sleep(100 * time.Millisecond)
	}
}

func getInfo(id string) (*avInfo, error) {
	url := fmt.Sprintf("https://api.bilibili.com/x/web-interface/archive/stat?aid=%v", id)
	request, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return nil, err
	}
	request.Header.Add("Accept", "*/*")
	request.Header.Add("Accept-Encoding", "gzip, deflate, br")
	request.Header.Add("Accept-Language", "zh-CN,zh;q=0.9")
	request.Header.Add("Connection", "keep-alive")
	request.Header.Add("Host", "api.bilibili.com")
	request.Header.Add("Origin", "https://www.bilibili.com")
	request.Header.Add("Referer", fmt.Sprintf("https://www.bilibili.com/video/av%v", id))
	request.Header.Add("Sec-Fetch-Mode", "cors")
	request.Header.Add("Sec-Fetch-Site", "same-site")
	request.Header.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.79 Safari/537.36")
	client := http.Client{}
	
	resp, err := client.Do(request)
	if resp.StatusCode != 200 {
		return nil, errors.New("Status Code Not 200")
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
		AID      int    `json:"aid"`
		Danmu    int    `json:"danmaku"`
		View     int    `json:"view"`
		Like     int    `json:"like"`
		Coin     int    `json:"coin"`
		Favorite int    `json:"favorite"`
	} `json:"data"`
}
