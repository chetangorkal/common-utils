package me.util;

import java.io.FileNotFoundException;
import java.io.FileReader;

import com.google.gson.Gson;
import com.google.gson.stream.JsonReader;

import me.jnvc.alumni.bean.TrelloCard;
import me.jnvc.alumni.bean.TrelloCards;

/**
 * Class to convert JSON of Trello card to text
 * 
 * @author Chetan Gorkal
 *
 */
public class TrelloCardJsonToTextConvertor {
	public static void main(String[] args) throws FileNotFoundException {

		Gson gson = new Gson();
		JsonReader reader = new JsonReader(new FileReader("trello_board.json"));
		TrelloCards data = gson.fromJson(reader, TrelloCards.class);
		for (TrelloCard card : data.getCards()) {
			System.out.println(card.getName());
		}
	}

}