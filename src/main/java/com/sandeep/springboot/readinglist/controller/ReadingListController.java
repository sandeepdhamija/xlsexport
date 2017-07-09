package com.sandeep.springboot.readinglist.controller;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.sandeep.springboot.readinglist.model.Book;
import com.sandeep.springboot.readinglist.repository.ReadingListRepository;
import com.sandeep.springboot.util.XLSWriter;

@Controller
@RequestMapping("/")
public class ReadingListController {
	private ReadingListRepository readingListRepository;

	@Autowired
	public ReadingListController(ReadingListRepository readingListRepository) {
		this.readingListRepository = readingListRepository;
	}
	
	@RequestMapping(value = "/export", method = RequestMethod.GET)
	public ResponseEntity readersExport() {
		String[] col={"first","second","third"};
	//	ByteArrayOutputStream resource=null;
		ByteArrayResource resource=null;
    try {
    	 resource=XLSWriter.writeObjectToXLS(null, col);
    	 // c= new ByteArrayResource(resource.toByteArray());
	} catch (IOException e) {
		// TODO Auto-generated catch block
		e.printStackTrace();
	}
    return ResponseEntity.ok().header("Content-Disposition", "attachment; filename=PreallocationData.xlsx")
			.contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
			.body(resource);
		/*List<Book> readingList = readingListRepository.findByReader(reader);
		if (readingList != null) {
			model.addAttribute("books", readingList);
		}*/
		//return "readingList";
	}

	@RequestMapping(value = "/{reader}", method = RequestMethod.GET)
	public String readersBooks(@PathVariable("reader") String reader, Model model) {
		List<Book> readingList = readingListRepository.findByReader(reader);
		if (readingList != null) {
			model.addAttribute("books", readingList);
		}
		return "readingList";
	}

	@RequestMapping(value = "/{reader}", method = RequestMethod.POST)
	public String addToReadingList(@PathVariable("reader") String reader, Book book) {
		book.setReader(reader);
		readingListRepository.save(book);
		return "redirect:/{reader}";
	}
}
