package com.sandeep.springboot.readinglist.repository;

import java.util.List;

import org.springframework.data.jpa.repository.JpaRepository;

import com.sandeep.springboot.readinglist.model.Book;

public interface ReadingListRepository extends JpaRepository<Book, Long> {

	List<Book> findByReader(String reader);
}
