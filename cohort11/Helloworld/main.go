package main

import "fmt"

type person struct {
	name   string
	age    int
	height float64
}

func main() {
	var p1 person
	p1 = person{
		name:   "Oreva",
		age:    23,
		height: 1.56,
	}

	p2 := person{
		name: "james",
		age: 25,
		height: 2.00,
	}
	fmt.Println(p1)
	fmt.Println(p2)
	fmt.Println(p1.age)
}
